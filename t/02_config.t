
# $Id: 02_config.t,v 1.5 2006/04/01 23:18:47 Daddy Exp $

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
if ($^O !~ m!win32!i)
  {
  diag(q'this is not Windows');
  exit 0;
  } # if
my $sMsg = $oICS->read || '';
my $iNoIIS = ($sMsg =~ m!can not find adsutil!i);
SKIP:
  {
  skip 'IIS is not installed on this machine?', 1 if $iNoIIS;
  isa_ok($object, 'Win32::IIS::Admin');
  $object->_parse_config;
  } # end of SKIP block

__END__

