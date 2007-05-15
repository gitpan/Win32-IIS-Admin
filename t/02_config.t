
# $Id: 02_config.t,v 1.9 2007/05/15 00:35:03 Daddy Exp $

use strict;
use warnings;

use ExtUtils::testlib;
use Test::More 'no_plan';
use IO::Capture::Stderr;

BEGIN
 {
 use_ok( 'Win32::IIS::Admin' );
 } # end of BEGIN block

my $oICS = new IO::Capture::Stderr;
$oICS->start;
my $o = Win32::IIS::Admin->new ();
$oICS->stop;
if ($^O !~ m!win32!i)
  {
  diag(q'this is not Windows');
  exit 0;
  } # if
my $sMsg = join(';', $oICS->read) || '';
my $iNoIIS = ($sMsg =~ m!can not find adsutil!i);
SKIP:
  {
  skip 'IIS is not installed on this machine?', 2 if $iNoIIS;
  isa_ok($o, 'Win32::IIS::Admin');
  my $sVersion = $o->iis_version;
  ok($sVersion);
  diag(qq{reported to be IIS version $sVersion});
  } # end of SKIP block

__END__

