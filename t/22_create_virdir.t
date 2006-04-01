
# $Id: 22_create_virdir.t,v 1.5 2006/04/01 23:18:47 Daddy Exp $

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
  skip 'IIS is not installed on this machine?', 6 if $iNoIIS;
  isa_ok($object, 'Win32::IIS::Admin');
  my $sDir = 'QQQperl_WIA_testQQQ';
  my $sRes = $object->_execute_script('adsutil', 'delete', "/W3SVC/1/Root/$sDir");
  like($sRes, qr'could not be found', 'iis does not already contain the path we want to add');
  like($sRes, qr"$sDir");
  $sRes = $object->create_virtual_dir(
                                      -dir_name => $sDir,
                                      -path => 'C:\doesnt\matter\whats\here',
                                      -executable => 1,
                                     );
  like($sRes, qr'Done\.', 'virtual dir created');
  $sRes = $object->_execute_script('adsutil', 'delete', "/W3SVC/1/Root/$sDir");
  like($sRes, qr'deleted path', 'virtual dir deleted');
  like($sRes, qr"$sDir");
  } # end of SKIP block

__END__
