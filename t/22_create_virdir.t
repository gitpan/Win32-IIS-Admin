
# $Id: 22_create_virdir.t,v 1.2 2006/03/19 03:43:08 Daddy Exp $

use ExtUtils::testlib;
use Test::More 'no_plan';

BEGIN
 {
 use_ok( 'Win32::IIS::Admin' );
 } # end of BEGIN block

my $object = Win32::IIS::Admin->new();
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

__END__
