
# $Id: Makefile.PL,v 1.2 2006/03/19 03:43:59 Daddy Exp $

use inc::Module::Install;

name('Win32-IIS-Admin');
version_from('lib/Win32/IIS/Admin.pm');
author('Martin Thurn (mthurn@cpan.org)');
abstract('Administer Internet Information Service on Windows');
license('perl');

build_requires('Test::Simple' => 0.44);
requires('Data::Dumper');
requires('File::Spec');
requires('Win32API::File');

&WriteAll;

__END__
