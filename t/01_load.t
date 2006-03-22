
# $Id: 01_load.t,v 1.4 2006/03/21 01:38:18 Daddy Exp $

use ExtUtils::testlib;
use Test::More 'no_plan';
use IO::Capture::Stderr;
use strict;
use warnings;

BEGIN
 {
 use_ok( 'Win32::IIS::Admin' );
 } # end of BEGIN block

my $oICS = new IO::Capture::Stderr;
$oICS->start;
my $object = Win32::IIS::Admin->new ();
$oICS->stop;
my $sMsg = $oICS->read || '';
my $iNoIIS = ($sMsg =~ m!can not find adsutil!i);
SKIP:
  {
  skip 'IIS is not installed on this machine?', 1 if $iNoIIS;
  isa_ok($object, 'Win32::IIS::Admin');
  } # end of SKIP block


__END__

