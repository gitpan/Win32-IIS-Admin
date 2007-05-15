
# $Id: Makefile.PL,v 1.5 2007/05/13 14:31:48 Daddy Exp $

use inc::Module::Install;

name('Win32-IIS-Admin');
version_from('lib/Win32/IIS/Admin.pm');
author('Martin Thurn (mthurn@cpan.org)');
abstract('Administer IIS webserver on Windows');
license('perl');

requires('Data::Dumper');
requires('File::Spec');
test_requires('IO::Capture::Stderr');
requires('IO::String');
test_requires('Test::Simple' => 0.44);
test_requires('Test::More');
recommends('Win32API::File');

&WriteAll;

__END__
