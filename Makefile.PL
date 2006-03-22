
# $Id: Makefile.PL,v 1.3 2006/03/21 01:36:00 Daddy Exp $

use inc::Module::Install;

name('Win32-IIS-Admin');
version_from('lib/Win32/IIS/Admin.pm');
author('Martin Thurn (mthurn@cpan.org)');
abstract('Administer IIS webserver on Windows');
license('perl');

build_requires('IO::Capture::Stderr');
requires('Data::Dumper');
requires('File::Spec');
build_requires('Test::Simple' => 0.44);
requires('Win32API::File');

&WriteAll;

__END__
