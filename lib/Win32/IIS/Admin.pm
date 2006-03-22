
# $Id: Admin.pm,v 1.5 2006/03/21 01:35:06 Daddy Exp $

=head1 NAME

Win32::IIS::Admin - Administer Internet Information Service on Windows

=head1 SYNOPSIS

  use Win32::IIS::Admin;
  my $oWIA = new Win32::IIS::Admin;
  $oWIA->create_virtual_dir(-dir_name => 'cgi-bin',
                            -path => 'C:\wwwroot\cgi-bin',
                            -executable => 1);

=head1 DESCRIPTION

Enables you to do a few administration tasks on a IIS webserver.
Currently only works for IIS 5 (i.e. Windows 2000 Server).
Currently there are very few tasks it can do.

=head1 METHODS

=over

=cut

package Win32::IIS::Admin;

use Data::Dumper;
use File::Spec::Functions;
use Win32API::File qw( :DRIVE_ );

use strict;

use constant DEBUG => 0;
use constant DEBUG_EXEC => 0;

use vars qw( $VERSION );
$VERSION = do { my @r = (q$Revision: 1.5 $ =~ /\d+/g); sprintf "%d."."%03d" x $#r, @r };

=item new

Returns a new Win32::IIS::Admin object, or undef if there is any problem
(such as, IIS is not installed on the local machine!).

=cut

sub new
  {
  my ($class, %parameters) = @_;
  # Find out where IIS is installed.
  # Find the cscript executable:
  my (@asTry, $sCscript);
  push @asTry, catfile($ENV{windir}, 'system32', 'cscript.exe');
  foreach my $sTry (@asTry)
    {
    if (-f $sTry)
      {
      $sCscript = $sTry;
      last;
      } # if
    } # foreach
  DEBUG && print STDERR " DDD cscript is ==$sCscript==\n";
  if ($sCscript eq '')
    {
    warn "can not find executable cscript\n";
    return undef;
    } # if
  # Get a list of logical drives:
  my @asDrive = Win32API::File::getLogicalDrives();
  DEBUG && print STDERR " DDD logical drives are: @asDrive\n";
  # See which ones are hard drives:
  my @asHD;
  foreach my $sDrive (@asDrive)
    {
    my $sType = Win32API::File::GetDriveType($sDrive);
    push @asHD, $sDrive if ($sType eq DRIVE_FIXED);
    } # foreach
  DEBUG && print STDERR " DDD hard drives are: @asHD\n";
  # Find the adsutil.vbs script:
  my $sAdsutil = '';
  @asTry = ();
  # This is the default location, according to microsoft.com:
  push @asTry, catdir($ENV{windir}, qw( System32 Inetsrv AdminSamples ));
  # This is where I find it on my old IIS installation:
  push @asTry, map { catdir($_, qw( inetpub AdminScripts )) } @asHD;
  @asTry = map { catfile($_, 'adsutil.vbs') } @asTry;
  foreach my $sTry (@asTry)
    {
    if (-f $sTry)
      {
      $sAdsutil = $sTry;
      last;
      } # if
    } # foreach
  DEBUG && print STDERR " DDD adsutil is ==$sAdsutil==\n";
  if ($sAdsutil eq '')
    {
    warn "can not find adsutil.vbs\n";
    return undef;
    } # if
  # Now we have all the info we need to get started:
  my %hash = (
              adsutil => $sAdsutil,
              cscript => $sCscript,
             );
  my $self = bless (\%hash, ref ($class) || $class);
  return $self;
  } # new


sub _parse_config
  {
  my $self = shift;
  my $sConfig = $self->_execute_script('adsutil', 'enum_all');
  DEBUG && printf STDERR (" DDD adsutil returned %d characters\n", length($sConfig));
  DEBUG && print STDERR " DDD   here they are ==$sConfig==\n" if (length($sConfig) < 999);
  use IO::String;
  my $oIS = IO::String->new($sConfig);
  if (0)
    {
    # See if we can use Config::IniFiles to read the config:
    die unless eval "use Config::IniFiles";
    my $oCI = Config::IniFiles->new(-file => 'SampleData/adsutil-enum_all.txt'); # \*oIS);
    print STDERR Dumper($oCI);
    print STDERR Dumper(\@Config::IniFiles::errors);
    # Doesn't work.
    } # if 0
  # Parse it ourself:
  my $sSection = 'root';
 LINE_OF_CONFIG:
  while (my $sLine = <$oIS>)
    {
    chomp $sLine;
    # Ignore empty lines and all-whitespace lines:
    next LINE_OF_CONFIG unless $sLine =~ m!\S!;
    my @asValue;
    if ($sLine =~ m!\A\[(.+)\]\Z!)
      {
      $sSection = "root$1";
      # DEBUG && print STDERR " DDD start section =$sSection=\n";
      } # if [SECTION]
    elsif ($sLine =~ m!\A(\S+)\s+:\s+\((\S+)\)\s*(.+)\Z!)
      {
      my ($sProperty, $sType, $sValue) = ($1, $2, $3);
      if ($sType eq 'STRING')
        {
        # Protect backslashes, in case this value is a dir/file path:
        $sValue =~ s!\\!/!g;
        $sValue = eval $sValue;
        @asValue = ($sValue);
        } # if STRING
      elsif ($sType eq 'INTEGER')
        {
        $sValue = eval $sValue;
        @asValue = ($sValue);
        } # if INTEGER
      elsif ($sType eq 'EXPANDSZ')
        {
        # Protect backslashes, this value is a dir/file path:
        $sValue =~ s!\\!/!g;
        $sValue = eval $sValue;
        $sValue =~ s!%([^%]+)%!$ENV{$1}!g;
        @asValue = ($sValue);
        } # if INTEGER
      elsif ($sType eq 'BOOLEAN')
        {
        $sValue = ($sValue eq 'True');
        @asValue = ($sValue);
        }
      elsif ($sType eq 'LIST')
        {
        @asValue = ();
        if ($sValue =~ m!(\d+)\sItems!)
          {
          my $iCount = 0 + $1;
 ITEM_OF_LIST:
          for (1..$iCount)
            {
            my $sSubline = <$oIS>;
            if ($sSubline =~ m!\A\s+\042([^"]+)\042!) #
              {
              push @asValue, $1;
              } # if
            else
              {
              print STDERR " WWW list item does not look like string, in line ==$sLine==\n";
              }
            } # for ITEM_OF_LIST
          } # if
        else
          {
          print STDERR " WWW found LIST type but not item count at line ==$sLine==\n";
          next LINE_OF_CONFIG;
          }
        } # if LIST
      elsif ($sType eq 'MimeMapList')
        {
        my %hash;
        while ($sValue =~ m!"(\S+)"!g)
          {
          my ($sExt, $sType) = split(',', $1);
          $hash{$sExt} = $sType;
          } # while
        @asValue = (\%hash);
        }
      else
        {
        print STDERR " EEE unknown type =$sType=\n";
        }
      local $" = ',';
      # DEBUG && print STDERR " DDD section=$sSection= property=$sProperty= type=$sType= value=@asValue=\n";
      $self->_config_set_value($sSection, $sProperty, $sType, @asValue);
      } # if PropertyName : (TYPE) value
    else
      {
      DEBUG && print STDERR " WWW unparsable line ==$sLine==\n";
      }
    } # while LINE_OF_CONFIG
  # print STDERR Dumper(keys %{$self->{config}->{root}->{W3SVC}->{2}});
  # print STDERR Dumper($self->{config}->{root}->{W3SVC}->{1});
  # print STDERR Dumper($self->{config}->{root}->{MimeMap});
  # exit 88;
  } # _parse_config

# FOR INTERNAL USE ONLY.  THIS FUNCTION DOES _NOT_ CHANGE THE
# CONFIGURATION OF IIS, ONLY OUR INTERNAL DATA STRUCTURES!

sub _config_set_value
  {
  my $self = shift;
  local $" = ',';
  # DEBUG && print STDERR " DDD _config_set_value(@_)\n";
  # Required arg1 = section:
  my $sSection = shift || '';
  return unless ($sSection ne '');
  # Required arg2 = parameter name:
  my $sParameter = shift || '';
  return unless ($sParameter ne '');
  # Optional arg3 = parameter type (default is STRING):
  my $sType = shift || 'STRING';
  # Remaining arg(s) will be taken as the value(s) for this parameter.
  return unless @_;
  $self->{config} ||= {};
  my $dest = $self->{config};
  # Convert the Section to a hierarchical list:
  my @asSection = split('/', $sSection);
  foreach my $s (@asSection)
    {
    $dest->{$s} ||= {};
    $dest = $dest->{$s};
    } # foreach
  if ($sType eq 'LIST')
    {
    $dest->{$sParameter} = \@_;
    }
  else
    {
    $dest->{$sParameter} = shift;
    }
  } # _config_set_value

sub _config_get_value
  {
  my $self = shift;
  local $" = ',';
  # DEBUG && print STDERR " DDD _config_get_value(@_)\n";
  # Required arg1 = section:
  my $sSection = shift || '';
  return unless ($sSection ne '');
  # Required arg2 = parameter name:
  my $sParameter = shift || '';
  return unless ($sParameter ne '');
  return unless ref($self->{config});
  my $dest = $self->{config};
  # Convert the Section to a hierarchical list:
  my @asSection = split('/', $sSection);
  foreach my $s (@asSection)
    {
    return unless ref($dest->{$s});
    $dest = $dest->{$s};
    } # foreach
  return $dest->{$sParameter};
  } # _config_get_value

=item create_virtual_dir

Given the following named arguments, create a virtual directory on the
default #1 server on the local machine's IIS instance.

=over

=item -dir_name => 'virtual'

This is the virtual directory name as it will appear to your browsers.

=item -path => 'C:/local/path'

This is the full path the the actual location of the data files.

=item -executable => 1

Give this argument if your virtual directory holds executable programs.
Default is 0 (NOT executable).

=back

=cut

sub create_virtual_dir
  {
  my $self = shift;
  my %hArgs = @_;
  $hArgs{-dir_name} ||= '';
  if ($hArgs{-dir_name} eq '')
    {
    $self->add_error(qq(Argument -dir_name is required on create_virtual_dir.));
    return;
    } # if
  $hArgs{-path} ||= '';
  if ($hArgs{-path} eq '')
    {
    $self->add_error(qq(Argument -path is required on create_virtual_dir.));
    return;
    } # if
  $hArgs{-executable} ||= 0;
  # print STDERR Dumper(\%hArgs);
  # We cravenly refuse to modify anything but the default #1 webserver:
  my $sWebsite = 1;
  # First, see if a virtual directory with the same name is already
  # exists:
  my $sSection = join('/', '', 'W3SVC', $sWebsite, 'Root', $hArgs{-dir_name});
  my $sPath = $self->_config_get_value($sSection, 'Path') || '';
  my $sRes = '';
  if ($sPath ne '')
    {
    # There is already a virtual directory with that name.  Create a
    # sensible error message:
    if ($sPath ne $hArgs{-path})
      {
      $self->add_error(qq(There is already a virtual directory named '$hArgs{-dir_name}', but it points to $hArgs{-path}));
      return;
      } # if
    $self->add_error(qq(There is already a virtual directory named '$hArgs{-dir_name}' pointing to $hArgs{-path}));
    # Fall through and (try to) set the access rules.
    } # if
  else
    {
    # Virtual dir not there, create it:
    $sRes .= $self->_execute_script('mkwebdir',
                                   qq(-v "$hArgs{-dir_name}","$hArgs{-path}"),
                                   qq(-w $sWebsite),
                                   # qq(-c $sComputer),
                                  ) || '';
    if ($sRes =~ m!Error!)
      {
      $self->add_error($sRes);
      return;
      } # if
    } # else
  # Whether the dir was already defined or not, try to set permissions
  # as requested:
  if ($hArgs{-executable})
    {
    # For some reason, the argument to chaccess has no leading slash:
    $sSection =~ s!\A/!!;
    # Set accesses for execution:
    $sRes .= $self->_execute_script('chaccess',
                                    -a => $sSection,
                                    qw( +execute +read +script ),
                                   );
    } # if
  return $sRes;
  } # create_virtual_dir


=item errors

Method not implemented.
In the current version, error messages are printed to STDERR as they occur.

=cut

sub errors
  {
  } # errors

sub _add_error
  {
  my $self = shift;
  print STDERR "@_\n";
  } # add_error

sub _execute_script
  {
  my $self = shift;
  my $sVBS = shift;
  my $sCmd = join(' ', $self->{cscript}, '-nologo', $self->{adsutil}, @_);
  $sCmd =~ s!adsutil!$sVBS!;
  DEBUG_EXEC && print STDERR " DDD exec ==$sCmd==\n";
  return qx/$sCmd/;
  } # _execute_script

=back

=head1 BUGS

To report a bug, please use L<http://rt.cpan.org>.

=head1 AUTHOR

Martin Thurn C<mthurn@cpan.org>

=head1 COPYRIGHT

This program is free software; you can redistribute
it and/or modify it under the same terms as Perl itself.

The full text of the license can be found in the
LICENSE file included with this module.

=cut

1;

__END__
