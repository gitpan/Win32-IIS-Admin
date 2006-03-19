
# $Id: 01_load.t,v 1.2 2005/10/01 17:23:11 Daddy Exp $

use ExtUtils::testlib;
use Test::More 'no_plan';

BEGIN
 {
 use_ok( 'Win32::IIS::Admin' );
 }

my $object = Win32::IIS::Admin->new ();
isa_ok($object, 'Win32::IIS::Admin');

__END__

