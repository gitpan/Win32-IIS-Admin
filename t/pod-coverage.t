# $Id: pod-coverage.t,v 1.1 2007/05/01 00:37:02 Daddy Exp $
use lib 'lib';
use Test::More;
my
$VERSION = do { my @r = (q$Revision: 1.1 $ =~ /\d+/g); sprintf "%d."."%03d" x $#r, @r };
eval "use Test::Pod::Coverage 1.00";
plan skip_all => "Test::Pod::Coverage 1.00 required for testing POD coverage" if $@;
all_pod_coverage_ok();
__END__
