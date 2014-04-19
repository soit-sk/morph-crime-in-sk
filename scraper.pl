#!/usr/bin/perl
# Copyright 2011, 2014 Lubomir Rintel <lkundrak@v3.sk>

use strict;
use warnings;

use LWP::Simple;
use HTML::Parser;
use URI;
use File::Basename qw/basename/;
use XML::Simple;
use Database::DumpTruck;

binmode STDOUT, ':utf8';
# http://www.minv.sk/?statistika_kriminality_v_slovenskej_republike_za_rok_2013
my $root = 'http://www.minv.sk/?kriminalita_2014_xml';
my $dt = new Database::DumpTruck ({ dbname => 'data.sqlite', table => 'data' });

# Array utilities
sub array_add
{
	my ($a1, $a2) = @_;
	map { $a1->[$_] + ($a2->[$_] || 0) } (0..$#$a1);
}

sub array_sub
{
	my ($a1, $a2) = @_;
	map { $a1->[$_] - ($a2->[$_] || 0) } (0..$#$a1);
}

# Return only sanitized content, no styles
sub sane_row
{
	my $row = shift;

	my @cells = map { $_->{Data}{content} }
		(ref $row->{Cell} eq 'ARRAY' ? @{$row->{Cell}} : ($_->{Cell}));

	# Chuck trailing undefs
	pop @cells unless defined $cells[$#cells];

	return \@cells;
}

# Return data from a single worksheet
sub grok_sheet
{
	my $sheet = shift;

	my $dept_code = $sheet->{'ss:Name'};
	return if $dept_code eq 'SR'; # Aggregation
	my @rows = map { sane_row $_ } @{$sheet->{Table}{Row}};

	# Assertion check
	shift (@rows)->[0] =~ /Zdroj Ministerstvo/
		or die 'The front fell off';

	# Sanitize date span
	shift (@rows)->[0] =~ /ZA OBDOBIE\s+(\d+)\.(\d+)\.\s+-\s+(\d+)\.(\d+)\.(\d+)/
		or die 'A wave hit the ship';
	my ($month, $year) = (sprintf ('%02d', $2), $5);

	# Department
	shift (@rows)->[0] =~ /^(.daje za .*)/
		or die 'One in million';
	my $dept_human = $1;

	# Assertion check, nothing useful in this header
	join ('', map { $_ || '' } @{shift (@rows)}) =~ /daje o trestn.*daje o st/
		or die 'Nothing else there';

	# Header
	my (undef, @headings) = @{shift (@rows)};
	splice @headings, 2, 1; # meaningless %

	# Read the actual body
	my @result;
	my @cat_path = (undef, undef, undef, undef);
	while (my $row = shift @rows) {
		my ($category, @data) = @$row;
		next unless defined $category;
		next if $category =~ /SPOLU$/;
		next if $category =~ /^CELKOV/;
		splice @data, 2, 1; # meaningless %

		# Clean up trailing whitespace
		$category =~ s/\s*$//;

		# Toplevel aggregation/category
		if ($category =~ /^[[:upper:]\s+]+$/) {
			# Sanitize the name
			$category =~ s/\s+/ /g;
			$category =~ s/(.)(.*)/$1\L$2/;
			@cat_path[0..$#cat_path] = ($category, undef, undef, undef);
		} elsif ($category !~ /^\s+-/) {
			@cat_path[1..$#cat_path] = ($category, undef, undef);
		} elsif ($category =~ /^\s{2,4}-\s*(.+)/) {
			@cat_path[2..$#cat_path] = ($1, undef);
		} elsif ($category =~ /^\s{5,8}-\s*(.+)/) {
			@cat_path[3..$#cat_path] = ($1);
		} else {
			die "Could not grok category: '$category'"
		}

		# Fix data
		@data = map { $_ || 0 } @data;

		# Construct the real row
		my @res = map { $_ ? $_ : () } @cat_path;
		$res[3] = $res[3] || undef;

		push @result, [\@res, \@data];
	}

	# Fix up aggregations
	foreach my $level (1..3) {
		my @agg;
		foreach my $row (reverse @result) {
			my ($cat_path, $data) = @$row;

			if ($cat_path->[$level + 1]) {
				next;
			} elsif ($cat_path->[$level]) {
				@agg = array_add ($data, \@agg);
			} elsif ($cat_path->[$level - 1]) {
				next unless @agg;
				$cat_path->[$level] = '(ostatne)';
				@$data = array_sub ($data, \@agg);
				@agg = ();
			}
		}
	}

	# Enrich with common data
	unshift @headings, 'Year', 'Month', 'Month/Year', 'Department Code', 'Department',
		'Category', 'Sub-category', 'Sub-category 2', 'Sub-category 3';
	@result = map { [ $year, $month, $year.$month, $dept_code, $dept_human, @{$_->[0]}, @{$_->[1]} ] }
		@result;

	return (\@headings, \@result);
}

# Format document
sub grok_xml
{
	# Absolutize
	my $uri = new URI (shift)->abs ($root);

	use URI::Escape;
	my $base = uri_unescape (basename($uri));
	return unless ($base =~ /_1_/ or $base =~ /^1/);

	my @rows;
	my $xml = XMLin (get($uri));
	foreach my $sheet (@{$xml->{Worksheet}}) {
		next unless my ($headers, $data) = grok_sheet ($sheet);
		s/[^a-zA-Z_0-9]/_/g foreach @$headers;
		foreach my $row (@$data) {
			push @rows, {};
			@{$rows[$#rows]}{@$headers} = @$row;
		}
	}

	$dt->insert (\@rows);
}

# Roll!
my $parser = new HTML::Parser (api_version => 3);
$parser->handler (start => sub {
	my $tag = shift;
	my $attr = shift;

	return unless $tag eq 'a';
	return unless exists $attr->{href};
	return unless $attr->{href} =~ m{/swift_data/source/policia/statistiky/.*.xml};
	grok_xml ($attr->{href});
}, 'tag, attr');

$parser->parse (get $root or die "I get it, but I don't get it. -- Alice Cooper");
$parser->eof;
