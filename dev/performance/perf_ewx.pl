#!/usr/bin/perl

##############################################################################
#
# Simple Perl program to test the speed and memory usage of the
# Excel::Writer::XLSX module.
#
# perl perf_ewx.pl [num_rows] [optimization_mode]
#
# Copyright 2013-2016, John McNamara, jmcnamara@cpan.org

use strict;
use warnings;
use Excel::Writer::XLSX;
use Time::HiRes qw(gettimeofday tv_interval);
use Devel::Size qw(total_size);

# Default to 1000 rows and non-optimised.
my $row_max  = $ARGV[0] || 1000;
my $col_max  = 50;
my $optimise = $ARGV[1] || 0;

# We double the rows below.
$row_max /= 2;

# Start timing after everything is loaded.
my $start_time = [gettimeofday];

# Start of program being tested.
my $workbook = Excel::Writer::XLSX->new( 'pl_ewx.xlsx' );

if ( $optimise ) {
    $workbook->set_optimization();
}

my $worksheet = $workbook->add_worksheet();

$worksheet->set_column( 0, $col_max, 18 );


for my $row ( 0 .. $row_max -1) {
    for my $col ( 0 .. $col_max -1 ) {
        $worksheet->write( $row * 2, $col, "Row: $row Col: $col" );
    }
    for my $col ( 0 .. $col_max ) {
        $worksheet->write( $row * 2 + 1, $col, $row + $col );
    }
}

# Get total memory size for workbook object before closing it.
my $total_size = total_size( $workbook );

$workbook->close();

# Get the elapsed time.
my $elapsed = tv_interval( $start_time );

# Print a simple CSV output for reporting.
printf "%6d, %3d, %6.2f, %d\n", $row_max * 2, $col_max, $elapsed, $total_size;


__END__
