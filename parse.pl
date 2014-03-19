#!/usr/bin/perl

use strict;
use warnings;
use HTML::TableExtract;
use Excel::Writer::XLSX;

my $workbook = Excel::Writer::XLSX->new('result.xlsx');
my $ws0      = $workbook->add_worksheet('PS_RF_STATS');
my $ws1      = $workbook->add_worksheet('PS_AUDIO_STATS');
$ws0->set_column( 0, 0, 20 );
$ws0->set_column( 1, 1, 21 );
$ws0->set_column( 2, 2, 20 );
$ws1->set_column( 0, 0, 20 );
$ws1->set_column( 1, 1, 21 );
$ws1->set_column( 2, 2, 20 );

my $t0 = new HTML::TableExtract->new( depth => 0, count => 0 );
$t0->parse_file('Mesa_19_feb_LU_m30dBm.html');
my $t1 = new HTML::TableExtract->new( depth => 0, count => 1 );
$t1->parse_file('Mesa_19_feb_LU_m30dBm.html');
my $t2 = new HTML::TableExtract->new( depth => 0, count => 2 );
$t2->parse_file('Mesa_19_feb_LU_m30dBm.html');
my $t3 = new HTML::TableExtract->new( depth => 0, count => 3 );
$t3->parse_file('Mesa_19_feb_LU_m30dBm.html');

my $n          = 2;

foreach my $row ( $t0->rows ) {
    $ws0->write( $n, 0, $row );
    $n++;
}

my $m = 2;

foreach my $row ( $t1->rows ) {
    $ws1->write( $m, 0, $row );
    $m++;
}

$n = $n + 5;
$m = $m + 5;

foreach my $row ( $t0->rows ) {
    $n++;
    $ws0->write( $n, 0, $row );
}

foreach my $row ( $t1->rows ) {
    $m++;
    $ws1->write( $m, 0, $row );
}
my $format = $workbook->add_format();
for my $i ( 0 .. 67 ) {
    $ws0->write_blank( $i, 3, $format );
}
for my $x ( 0 .. 47 ) {
    $ws1->write_blank( $x, 3, $format );
}

open( my $in, "<", "Mesa_19_feb_LU_m30dBm.html" )
  or die "Can't open Mesa_19_feb_LU_m30dBm.html: $!";

my $substring = "";
my @USBID     = "";
my @dt        = "";

while (<$in>) {

    if (m /<h1>EHIF transaction/s) {

        push @USBID, substr $_, 40,  6;
        push @dt,    substr $_, 237, 23;
        push @dt,    substr $_, 286, 23;
        push @dt,    substr $_, 321, 23;
        push @dt,    substr $_, 240, 23;
        push @dt,    substr $_, 289, 23;
        push @dt,    substr $_, 324, 23;
    }
}

$ws0->write_string( 'A1', 'USB ID:' );
$ws0->write_string( 'B1', $USBID[1] );

$ws0->write_string( 'A37', 'USB ID:' );
$ws0->write_string( 'B37', $USBID[3] );

$ws1->write_string( 'A1', 'USB ID:' );
$ws1->write_string( 'B1', $USBID[5] );

$ws1->write_string( 'A27', 'USB ID:' );
$ws1->write_string( 'B27', $USBID[7] );

$ws0->write_string( 'A33', 'Transaction started:' );
$ws0->write_string( 'B33', $dt[1] );

$ws0->write_string( 'A34', 'Transaction finished:' );
$ws0->write_string( 'B34', $dt[2] );

$ws0->write_string( 'A35', 'Logged:' );
$ws0->write_string( 'B35', $dt[3] );

$ws1->write_string( 'A23', 'Transaction started:' );
$ws1->write_string( 'B23', $dt[16] );

$ws1->write_string( 'A24', 'Transaction finished:' );
$ws1->write_string( 'B24', $dt[17] );

$ws1->write_string( 'A25', 'Logged:' );
$ws1->write_string( 'B25', $dt[18] );

$ws0->write_string( 'A69', 'Transaction started:' );
$ws0->write_string( 'B69', $dt[25] );

$ws0->write_string( 'A70', 'Transaction finished:' );
$ws0->write_string( 'B70', $dt[26] );

$ws0->write_string( 'A71', 'Logged:' );
$ws0->write_string( 'B71', $dt[27] );

$ws1->write_string( 'A49', 'Transaction started:' );
$ws1->write_string( 'B49', $dt[41] );

$ws1->write_string( 'A50', 'Transaction finished:' );
$ws1->write_string( 'B50', $dt[42] );

$ws1->write_string( 'A51', 'Logged:' );
$ws1->write_string( 'B51', $dt[43] );

