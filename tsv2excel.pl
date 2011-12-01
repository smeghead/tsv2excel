#!/usr/bin/perl -w
use strict;
use warnings;
use utf8;

use Spreadsheet::WriteExcel;
use Data::Dumper;

sub usage {
    print STDERR "usage: \n";
    print STDERR " cat <csv filename> | $0 <output filename>\n";
    exit -1;
}

my $output_filename = shift;
my $csv_content = do {
    local $/;
    <STDIN>;
};

print Dumper($output_filename);
print Dumper($csv_content);

usage() unless $output_filename;
usage() unless $csv_content;

my $storage_book = {};

#prepare data
my $row_count = 0;
my $rows = {};
for my $line (split /\n/, $csv_content) {
    my $col_count = 0;
    for my $field (split /,\s*/, $line) {
        $rows->{$row_count}->{$col_count++} = $field;
    }
    $row_count++;
}
$storage_book->{click} = $rows;

my $dest_book  = Spreadsheet::WriteExcel->new("$output_filename")
    or die "Could not create a new Excel file in $output_filename: $!";
print "\n\nSaving recognized data in $output_filename...";
foreach my $sheet (keys %$storage_book) {
    my $dest_sheet = $dest_book->addworksheet($sheet);
    foreach my $row (keys %{$storage_book->{$sheet}}) {
        foreach my $col (keys %{$storage_book->{$sheet}->{$row}}) {
            $dest_sheet->write($row, $col, $storage_book->{$sheet}->{$row}->{$col});
        }
    }
}
$dest_book->close();
print " done!\n";
