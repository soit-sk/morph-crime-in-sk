#!/usr/bin/perl
# Public Domain: Can be used, modified and distributed without any restriction
# Copyright 2011, 2014, 2015 Lubomir Rintel <lkundrak@v3.sk>

use strict;
use warnings;

use LWP::Simple;
use HTML::Parser;
use URI;
use File::Basename qw/basename/;
use XML::Simple;
use Database::DumpTruck;

binmode STDOUT, ':utf8';
my $dt = new Database::DumpTruck ({ dbname => 'data.sqlite', table => 'data' });

# Uniquely identifies a category
my @id = ('Department Code', 'Department', 'Category', 'Sub-category', 'Sub-category 2', 'Sub-category 3');

# Uniquely identifies an database entry (usable in an unique index)
my @keys = ('Year', 'Month', 'Month_Year', @id);

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
	return {} if $dept_code eq 'SR'; # Aggregation
	my @rows = map { sane_row $_ } @{$sheet->{Table}{Row}};

	# Assertion check
	shift (@rows)->[0] =~ /Zdroj Ministerstvo/
		or die 'The front fell off';

	# Sanitize date span
	shift (@rows)->[0] =~ /ZA OBDOBIE\s+(\d+)\.(\d+)\.\s+-\s+(\d+)\.(\d+)\.(\d+)/
		or die 'A wave hit the ship';
	my ($month, $year) = (sprintf ('%02d', $4), $5);

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
	my %result;
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
			@cat_path[0..$#cat_path] = ($category, '', '', '');
		} elsif ($category !~ /^\s+-/) {
			@cat_path[1..$#cat_path] = ($category, '', '');
		} elsif ($category =~ /^\s{2,4}-\s*(.+)/) {
			@cat_path[2..$#cat_path] = ($1, '');
		} elsif ($category =~ /^\s{5,8}-\s*(.+)/) {
			@cat_path[3..$#cat_path] = ($1);
		} else {
			die "Could not grok category: '$category'"
		}

		# Fix data
		@data = map { $_ || 0 } @data;

		# Construct the real row
		#my @res = map { defined $_ ? $_ : () } @cat_path;
		#$res[3] = $res[3] || undef;

		my %row;
		@row{@keys, @headings} = ($year, $month, $year.$month, $dept_code, $dept_human, @cat_path, @data);
		my $key = join "|", map { $row{$_} } @id;
		$result{$key} = \%row;
	}

	return \%result;
}

# Format document
sub grok_xml
{
	# Absolutize
	my $uri = shift;
	my %month;

	use URI::Escape;
	my $base = uri_unescape (basename($uri));
	return unless ($base =~ /_1_/ or $base =~ /^1/);

	my @rows;
	my $xml = XMLin (get($uri));
	%month = (%month, %{grok_sheet ($_)}) foreach @{$xml->{Worksheet}};

	return \%month;
}

sub month_sub
{
	my $a = shift;
	my $b = shift;

	return $a unless $b;

	foreach my $id (keys %$a) {
		foreach my $key (keys %{$a->{$id}}) {
			next if grep { $_ eq $key } @keys;
			$a->{$id}{$key} -= $b->{$id}{$key};
		}
	}

	return $a;
}

sub process_year
{
	my $uri = shift;

	my $previous_month;
	my $parser = new HTML::Parser (api_version => 3);
	$parser->handler (start => sub {
		my $tag = shift;
		my $attr = shift;

		return unless $tag eq 'a';
		return unless exists $attr->{href};
		return unless $attr->{href} =~ m{/swift_data/source/policia/statistiky.*.xml};
		my $this_month = grok_xml (new URI ($attr->{href})->abs ($uri)) or return;

		$this_month = month_sub ($this_month, $previous_month);
		$previous_month = $this_month;

		$dt->upsert ([values %$this_month]);
	}, 'tag, attr');

	$parser->parse (get $uri or die "I get it, but I don't get it. -- Alice Cooper");
	$parser->eof;

}

my $root = 'http://www.minv.sk/?statistika-kriminality-v-slovenskej-republike-xml';

# Roll!
my $parser = new HTML::Parser (api_version => 3);
$parser->unbroken_text (1);
$parser->handler (start => sub {
	my $tag = shift;
	my $attr = shift;

	return unless $tag eq 'a';
	return unless exists $attr->{href};
	return unless $attr->{href} =~ m{/.*za.rok.(\d\d\d\d)$};

	process_year (new URI ($attr->{href})->abs ($root));
}, 'tag, attr');

$parser->parse (get $root or die);
$parser->eof;

# Ensure there's indices when we know there's a scheme
$dt->create_index (\@keys, undef, 1, 1);
