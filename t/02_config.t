
# $Id: 02_config.t,v 1.2 2005/10/01 17:22:10 Daddy Exp $

use ExtUtils::testlib;
use Test::More 'no_plan';

BEGIN
 {
 use_ok( 'Win32::IIS::Admin' );
 } # end of BEGIN block

my $object = Win32::IIS::Admin->new ();
isa_ok($object, 'Win32::IIS::Admin');

$object->_parse_config;

__END__

