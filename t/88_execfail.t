
# $Id: 88_execfail.t,v 1.1 2006/03/19 03:42:45 Daddy Exp $

use ExtUtils::testlib;
use Test::More 'no_plan';

BEGIN
 {
 use_ok( 'Win32::IIS::Admin' );
 use_ok( 'IO::Capture::Stderr' );
 } # end of BEGIN block

my $oICS = new IO::Capture::Stderr;
my $object = Win32::IIS::Admin->new();
isa_ok($object, 'Win32::IIS::Admin');

$oICS->start;
print STDERR $object->_execute_script('adsutil', 'totally wrong args');
$oICS->stop;
like($oICS->read, qr(Command not recognized));
$oICS->start;
print STDERR $object->_execute_script('chaccess', '-a /W3SVC/1/no_such_path +read +execute');
$oICS->stop;
my $sMsg = $oICS->read;
like($sMsg, qr(80005000));
like($sMsg, qr(Unable to open specified node));

__END__
