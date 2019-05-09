#!/usr/bin/perl
#
# Name: DonorImport 1.3.02
# Date: -5-09-2019
# Copyright: 2019 Michael Bates
# Website: https://bates.link/donorimport/
# Author: Michael Bates
# Testing: Dan Guenther
# This program is a derivative work based the code of DonorImport 1.1.4 (Jan. 20, 2013) developed by Jordan Boland and Dan Guenther
# It is released under the GPL Public License
#

use strict;
use warnings;


# MD5 is a type of cryptographic hash function.  Although it is no longer secure for use with passwords
# it will still work perfectly well for our purposes (creating unique keys for database inserts).
# The chance of two entries creating the same hash value is one in (unreasonably large number).
use Digest::MD5 qw( md5_hex );


# Constants
use constant DONORIMPORT_VERSION => "1.3.02";

# constants defining the index of our program's record digest
use constant {
	DIG_KEEP_RECORD				=> 0,
	DIG_RECORD_NUM				=> 1,
	DIG_DONOR_ID				=> 2,
	DIG_DONOR_NAME				=> 3,
	DIG_DONOR_ADDR				=> 4,
	DIG_DONOR_CITY				=> 5,
	DIG_DONOR_STATE				=> 6,
	DIG_DONOR_ZIP				=> 7,
	DIG_OFFERING_DATE			=> 8,
	DIG_OFFERING_AMOUNT			=> 9,
	DIG_DONOR_FIRSTNAME			=> 10,
	DIG_DONOR_LASTNAME_ORG		=> 11,
	DIG_RECEIPT_ID				=> 12,
	DIG_PERSON_TYPE				=> 13,
	DIG_ORG_CONTACT_PERS		=> 14,
	DIG_GIFT_MEMO				=> 15,
	DIG_DONOR_MI				=> 16,
	DIG_DONOR_SFFX				=> 17,
	DIG_SP_FIRST				=> 18,
	DIG_SP_MI					=> 19,
	DIG_SP_LAST					=> 20,
	DIG_PHONE					=> 21,
	DIG_CLASS					=> 22,
};


# Global variables
#   @records is an array to store a digest of our records
use vars qw(@records $monthly_OK $dialog $title $max_width);
$monthly_OK = 0; # sets user confirmation that monthly file processing is okay
$dialog = ''; # dialog for user prompts
$title = '';  # title for user prompts
$max_width = 70; # max width for wordwrap

# Create a hash to store classes which represent donations.
# Classes not on this list will be treated as internal AGWM accounting transactions -- not donations
# Also, the values are used as descriptions for the gift memo field
my %donation_classes = (
	'0' 	=> '(00) Work Support',
	'00'	=> '(00) Work Support',
	'9' 	=> '(09) Special personal',
	'09'	=> '(09) Special personal',
	'40'	=> '(40) Work (Institutional)',
	'42'	=> '(42) Bible School',
	'44'	=> '(44) Radio-TV-Films',
	'46'	=> '(46) Relief',
	'50'	=> '(50) National Worker',
	'60'	=> '(60) Building',
	'62'	=> '(62) Building',
	'64'	=> '(64) Building'
	);


# SUBROUTINES


# Declare trim helper function.  Remove leading and tailing spaces, and quotes from fields
sub trim($) {
    my $string = shift;
    $string =~ s/"//g;
	$string =~ s/^\s+//;
	$string =~ s/\s+$//;
	return $string;
}

# Detects if is running under MacOS
#
# @return int 1     If MacOS
# @return int 0     If not MacOS
sub is_mac() {
    if ($^O eq 'darwin' or $^O eq 'MacOS' or $^O eq 'rhapsody') {
        return 1;
    }
    else {
        return 0;
    }
}

# Word-wrap function
# Wraps text by adding new Lines
# global $max_width overrides the default width limit inside this sub
#
# @param string $paragraph  - paragraph to wordwrap
#
# @return string $paragraph - wordwrapped paragraph
sub wordwrap {
	my $paragraph = shift;

    # Set default $limit
    my $limit = 70;
    # Use value of global $max_width if defined
    if( defined($max_width) && $max_width > 10 ) {
        $limit = $max_width;
    }

	$paragraph =~ s/([^\n]{0,$limit})(?:\b\s*|\n)/$1\n/gio;
	return $paragraph;
}

# MacOS dialogs using AppleScript
# Be aware that dialog strings with single-quotes must be escaped for AppleScript

# MacOS Y/N dialog
sub dialog_yn {
  my ($text) = ( @_ > 0 ) ? shift : '';
  my ($title) = ( @_ > 0 ) ? shift : '';
  return `osascript -e '
        tell app "System Events"
          button returned of (display dialog "$text" & return buttons {"Yes", "No"} default button "Yes" with title "$title")
        end tell'
  `;
}

# MacOS OK dialog
sub dialog_ok {
  my ($text) = ( @_ > 0 ) ? shift : '';
  my ($title) = ( @_ > 0 ) ? shift : '';
  return `osascript -e '
        tell app "System Events"
          button returned of (display dialog "$text" & return buttons {"OK"} default button "OK" with title "$title")
        end tell'
  `;
}


# MacOS exit dialog
sub dialog_exit {
  my ($text) = ( @_ > 0 ) ? shift : '';
  my ($title) = ( @_ > 0 ) ? shift : '';
	`osascript -e '
        tell app "System Events"
          button returned of (display dialog "$text" & return buttons {"OK"} default button "OK" with title "$title")
        end tell'
  `;
  exit;
}


# Windows and non-MacOS dialogs...
# Waits for keypress, then exits (Windows)
sub key_exit() {
	use Term::ReadKey;
	print "\nPress any key to exit...\n";
	ReadMode('cbreak');
	my $key = ReadKey(0);
	ReadMode('normal');
	exit;
}

# Waits for keypress, then continues (Windows)
sub key_continue() {
	use Term::ReadKey;
	print "\nPress any key to continue...\n";
	ReadMode('cbreak');
	my $key = ReadKey(0);
	ReadMode('normal');
}

# Reads one key, returns $key (Windows)
sub get_key() {
    use Term::ReadKey;
    ReadMode('cbreak');
    my $key = ReadKey(0);
    ReadMode('normal');
    return($key);
}



# Get filename without get_filename_without_extension
#
# Uses a regular expression to determine the filename without the extension.
# We will use this to remove '.csv' from the end of our output filename and
#
# @param string $filename   Input $filename
# @return string            The filename without the extension
sub get_filename_without_extension {
    my $filename = shift;
    if ($filename =~ /^(.*)\.[^.]+$/) {
        return $1;
    } else {
        return $filename; # in the case of no filename extension
    }
}




# parseNames
# Input FirstName, LastName (fields from CSV)
# Output: parsed names: $first,$MI,$last,$suffix,$spouseFirst,$spouseMI,$spouseLast
sub parseNames
{
		my $inputFirst = $_[0];
		my $inputLast = $_[1];
		my $spouseFirst = '';
		my $spouseMI = '';
		my $spouseLast = '';
		my $first = '';
		my $MI = '';
		my $last = '';
		my $suffix = '';

		if($inputFirst . $inputLast ne '') { # Has a name listed
			($first, $spouseFirst) = sepSpouseNames($inputFirst); # separate husband's and wife's names
			($first, $MI) = parseFirstMI($first); # separate first & MI
			($last, $suffix) = parseLastSuffix($inputLast);
			if ($spouseFirst ne '') {
				($spouseFirst, $spouseMI) = parseFirstMI($spouseFirst); # separate spouse first & MI
				$spouseLast = $last;
			}
		}
		else { # is blank
			($first, $MI, $last, $suffix) = ('','','','');
		}
		return ($first,$MI,$last,$suffix,$spouseFirst,$spouseMI,$spouseLast);
}



# Separate spouse names by '&'
# Dependencies: trim() function to remove leading, trailing whitespace
# Input: String with FirstName field from CSV (Ex. 'John H & Sue' )
# Ouput: $firstSpouse, $secondSpouse    (Ex. 'John H', 'Sue' )
sub sepSpouseNames($)
{
	my $input = $_[0];
	my $firstSpouse = '';
	my $secondSpouse = '';
	if ($input =~ m/(^[A-Za-z\s]{3,})\&([A-Za-z\s]{3,})$/){ # has a spouse name after '&' character
		$firstSpouse = $1;
		$secondSpouse = $2;
	}
	else {
		$firstSpouse = $input;
		$secondSpouse = '';
	}
	return (trim($firstSpouse), trim($secondSpouse));
}

# ParseFirstMI
# Parse a string that contains a First Name followed by a MI  (Ex. 'Michael R')
# Return FirstName and MI separately.
# If First Name is an initial followed by a longer Middle Name (Ex. 'W Russell'
#    treat it as a person who goes by his/her middle name and substitute Middle Name as FirstName, discarding first initial
# Input: String with FirstName field  (Note, should process with sepSpouseNames() first)
# Output: $firstName, $MI
sub parseFirstMI($)
{
	my $input = $_[0];
	my $first = '';
	my $MI = '';
	if ($input =~ m/(^[A-Za-z]{2,})\s([A-Za-z])$/) { # Match 2+ letter first name and 1 letter MI
		$first = $1;
		$MI = $2;
	}
	elsif ($input =~ m/^[A-Za-z]\s([A-Za-z]{3,}$)/) { # Match person who goes by middle name, like 'W Russell' and use 'Russell' as First
		$first = $1;
	}
	else { # otherwise, treat entire string as first name
		$first = $input;
	}
	return (trim($first),trim($MI));
}

# parseLastSuffix
# Input:  String with Last name possibly followed by suffix (Ex. Bates Jr)
# Output: ($last, $suffix)
sub parseLastSuffix($)
{
	my $input = $_[0];
	my $last = '';
	my $suffix = '';
	if ($input =~ m/(^[A-Za-z\s]{3,})\s(Jr|Sr|I|II|III)\.*$/i) {
		$last = $1;
		$suffix = $2;
	}
	else
	{
		$last = $input;
	}
	return (trim($last),trim($suffix));
}



# Get ledger with unique donations
#
# This subroutine receives a ledger array (@records) as well as hashes with
# references to the positive and negative sides of the ledger.
# It drills down into it using hashes to find unique donations. It will increment a counter
# when there are similar donations of the same amount and date from a specific
# donor, but in a separate class.
#
# @param reference \@records    - array of records
# @param reference  \%positive  - hash with positive ledger
# @param reference  \%negative  - hash with negative ledger
# @return array @records
sub get_ledger_with_unique_donations {
	# Dereference ledger
	my @records = @{shift()};

	# For each catalog
	foreach my $hashRef (@_) {

        # Drilling down - for each donor
		while (my ($donorID, $hashRefA) = each (%$hashRef)) {

            # Drilling down - for each donation date
			while (my ($offeringDate, $hashRefB) = each(%$hashRefA)) {

                # Drilling down - for each amount
				while (my ($amount, $hashRefC) = each(%$hashRefB)) {

					# This counter is used to count records in inner loop. It is reset to zero within the outer (same $amount) loop
					# because the MD5 hash function used for generating a unique donation id does not take class into account -mrb
					my $counter = 0;
					while (my ($class, $arrayRef) = each(%$hashRefC)) {

                        # make a copy of this array so we can iterate over it and keep track of the count of offerings
						my @tmp = @$arrayRef;

                        # Drilling down - for each donation (donation amounts possibly repeat)
						foreach my $record_number (0..$#tmp) {

                            # Twiddle
							$records[$tmp[$record_number]][DIG_KEEP_RECORD] = 1;
							$records[$tmp[$record_number]][DIG_RECORD_NUM] = $counter;
							$counter++;
						}
					}
				}
			}
		}
	}
	return @records; # Return the new ledger
}






# MAIN PROGRAM

#  Display help message if no files given as arguments
if (@ARGV == 0) {
	$dialog = "DonorImport version ".DONORIMPORT_VERSION."\n\n" .
              "No files passed to DonorImport.\n" .
              "To use DonorImport, please use your mouse to drag an AGWM CSV " .
              "file on top of the DonorImport icon. For more info, please ".
              "refer to the online videos and the support page at \n".
              "https://bates.link/donorimport/start-here/\n";
	$title = "Alert";
    print STDERR wordwrap($dialog);
	if (is_mac()) {
		$dialog = $dialog . "The program will exit now...";
		dialog_exit($dialog, $title);
	}
	else {
		key_exit();
	}
}

# loop through each input file. Using FILE as program block label
FILE: while(my $input_filename = shift(@ARGV)) {

    # Declare several variables within the scope of processing an individual file

	my @records;   #holds imported donation records
    my $gifts = '';
    my $donors = '';

    # filetype indicator: 'monthly' or 'daily'
    my $filetype = undef;

    # Mapping for CSV fields that we need to import
    # Fields common to both daily and monthly files
    my %csv_common_fields = (
                          DonorAcctNo => undef,
                          DonorLastName => undef,
                          DonorFirstName => undef,
                          DonorName => undef,
                          DonorStreet => undef,
                          DonorCity => undef,
                          DonorState => undef,
                          DonorZip => undef,
                          DonorTelephone => undef,
                          ClassNo => undef,
                         );

    # Fields only in daily files
    my %csv_daily_fields = (
                          ContributionAmount => undef,
                          ContributionDate => undef,
                          ReceiptNumber => undef,
                          Remarks => undef,
                        );

    # Fields only in monthly files
    my %csv_monthly_fields = (
                          CurrentDate => undef,
                          CurrentAmt => undef,
                          YTDAmt => undef,
                          );


    # Step 1. Open the file for input
    #  - Display message:  Processing file ...
	print "\nProcessing file (${input_filename})...\n";

    #  - First check that file exists and is readable
    #  - Then open for input, with FILEHANDLE $input_fh
    my $input_fh;
    if(-e $input_filename && -f _ && -r _ ) {
        open($input_fh, '<', $input_filename) or do {
                    print STDERR "Cannot open file $input_filename. Exiting...\n";
                    sleep(3);
                    die "$!";
                };
    } else {
        die "Cannot find or read '$input_filename'\n";
    }


    # Step 2. Process header fields and create map
    #  - Process header fields
    my $header = <$input_fh>;
    chomp $header;
    # Note: Not using Text::CSV here due to design requirement to not include external libraries on MacOS.
    #       Code below does not support fields with embedded commas.
    my @header_fields = split "," , $header;

    #  - Map CSV fields to our hast lists
    foreach my $i (0 .. $#header_fields) {
        # trim whitespace and quotes
        $header_fields[$i] = trim($header_fields[$i]);

        # Compare field name to our list of csv fields we wish to import
        # Map index to hash lists for common, daily or monthly fields
        # also, set an filetype indicator
        my $field_name = $header_fields[$i];
        if( exists($csv_common_fields{$field_name})  ) {
            $csv_common_fields{$field_name} = $i; # map to this index

        } elsif( exists($csv_daily_fields{$field_name})  ) {
            $csv_daily_fields{$field_name} = $i;
            $filetype = 'daily';

        } elsif ( exists($csv_monthly_fields{$field_name})  ) {
            $csv_monthly_fields{$field_name} = $i;
            $filetype = 'monthly';

        } else {
            next; # extra field that we don't need to import
        }
    }



    # Step 3. Confirm that we mapped all the required header $field_separator
    # We'll assume csv fields are valid until we find an undefined key

    # - First loop through fields common to monthly & daily files
    foreach my $key (keys(%csv_common_fields)) {
        if(!defined($csv_common_fields{$key})) {
            print STDERR "CSV file headers missing a required field: $key. Exiting...\n";
            sleep(5);
            die;
        }
    }

    # - If monthly, loop through monthly fields
    if ( defined($filetype) && $filetype eq 'monthly' ) {
        foreach my $key (keys(%csv_monthly_fields)) {
            if(!defined($csv_monthly_fields{$key})) {
                print STDERR "CSV file headers missing a required field: $key. Exiting...\n";
                sleep(5);
                die;
            }
        }
    }
    # If daily, loop through daily fields
    elsif ( defined($filetype) && $filetype eq 'daily' ) {
        foreach my $key (keys(%csv_daily_fields)) {
            if(!defined($csv_daily_fields{$key})) {
                print STDERR "CSV file headers missing a required field: $key. Exiting...\n";
                sleep(5);
                die;
            }
        }
    }
    # Otherwise, filetype is undetermined
    else {
        print STDERR "Unable to determine if this is a monthly or daily CSV file. \n";
        sleep(5);
        die;
    }



    # Give notice to user if this is a monthly file
    if ((!$monthly_OK) and ($filetype eq 'monthly')) {
        $dialog = "DonorImport has detected that this file is a monthly giving ".
            "summary. Please be aware that importing both daily and monthly CSV " .
            "files for the same month can result in duplicate donations in your " .
			"TntConnect database.\n";
        $title = "Alert";
        print wordwrap($dialog);
        my $key = '';

        if (is_mac()) { # MacOS - use AppleScript Y/N dialog
            $dialog = $dialog . "Do you wish to proceed?";
            $key = dialog_yn($dialog, $title);
            chomp $key;
            if ($key eq 'Yes') {
                # continue processing
            }
            else { # No or unrecognized response
                next FILE; # skip file
            }
        }
        else {  # Not MacOS. Use get_key()
            my $valid_response = 0;
            until ($valid_response) {
                print "Do you wish to proceed?  (Y)es, (N)o, Yes to (A)ll?\n";
                $key = get_key();
                if ($key eq 'A' or $key eq 'a') {
                    $monthly_OK = 1;
                    $valid_response = 1;
                }
                elsif ($key eq 'Y' or $key eq 'y') {
                    $valid_response = 1;
                }
                elsif ($key eq 'N' or $key eq 'n') {
                    $valid_response = 1;
                    next FILE;
                }
                else {
                    print "Unrecognized reponse. Please type Y, N, or A.\n";
                }
            }
        }
    }  # end monthly file user prompt confirmation



    # Read lines in file and store records in memory
    foreach my $line (<$input_fh>)
    {
        # Skip totals line
        next if $line =~ /totals >>/;

        # Split on commas. Field-embedded commas not supported!
        my @record = split(',', $line);

        # For each field in the record, clean up extra whitespace
        foreach my $i (0..$#record) {
            $record[$i] = trim($record[$i]);
        }

        my @tntmpdigrecord = (0, 0); # DIG_KEEP_RECORD, DIG_RECORD_NUM

        # Store needed field values in variables
        # Initiate them with empty values
        my ($acctNoLU,
            $donorNameFirstLU,
            $donorNameLastLU,
            $donorOrgNameLU,
            $streetLU,
            $cityLU,
            $stateLU,
            $zipLU,
            $phone,
            $classLU,
            $rcptID_or_YTDAmt,
            $contribDateLU,
            $contribAmtLU,
            $remarks
            ) = ('', '', '', '', '', '', '', '', '', '', '', '', '', '');

        # - First, store field values common to both monthly and daily files
        $acctNoLU           = $record[$csv_common_fields{DonorAcctNo}];
        $donorNameFirstLU   = $record[$csv_common_fields{DonorFirstName}];
        $donorNameLastLU    = $record[$csv_common_fields{DonorLastName}];
        $donorOrgNameLU     = $record[$csv_common_fields{DonorName}];
        $streetLU           = $record[$csv_common_fields{DonorStreet}];
        $cityLU             = $record[$csv_common_fields{DonorCity}];
        $stateLU            = $record[$csv_common_fields{DonorState}];
        $zipLU              = $record[$csv_common_fields{DonorZip}];
        $phone              = $record[$csv_common_fields{DonorTelephone}];
        $classLU            = $record[$csv_common_fields{ClassNo}];

        # - Then, if daily, store the fields only in daily files
        if($filetype eq 'daily') {
            $rcptID_or_YTDAmt  = $record[$csv_daily_fields{ReceiptNumber}];
            $contribDateLU     = $record[$csv_daily_fields{ContributionDate}];
            $contribAmtLU      = $record[$csv_daily_fields{ContributionAmount}];
            $remarks           = $record[$csv_daily_fields{Remarks}];
        }
        # - Otherwise, store field values in monthly files
        else {
            $rcptID_or_YTDAmt  = $record[$csv_monthly_fields{YTDAmt}];
            $contribDateLU     = $record[$csv_monthly_fields{CurrentDate}];
            $contribAmtLU      = $record[$csv_monthly_fields{CurrentAmt}];
        }


        # Determine if contribution and generate gift memo field
        #
        # First, it determines if this represents a donation based on if the
        # class number exists in the global %donation_classes hash.
        #
        # Secondly, it generates the gift memo field by looking up the
        # value describing that class. For example:
        #           '00' => '(00) Work support'
        #
        # It also appends any comments from the donor if present in the remarks field
        my $gift_memo = '';
        if (exists($donation_classes{$classLU})) {
            $gift_memo = $donation_classes{$classLU};
            if ( $remarks ne '' ) { # has a gift remark, then append
                $gift_memo = $gift_memo . ': ' . $remarks;
            }
        } else { # Not donation, skip this record
            next;
        }


        # Parse Names
        # First, declare some variables
        my $donorName = '';
        my $donorFirst = '';
        my $donorMI = '';
        my $donorLastOrg = '';
        my $donorSuffix = '';
        my $spFirst = '';
        my $spMI = '';
        my $spLast = '';
        my $orgContactPers = '';
        my $personType = ''; # Valid values: 'P', 'O'

        if ($donorOrgNameLU ne '') { # Has organization name, is organization
            if ($donorNameFirstLU.$donorNameLastLU ne '') { # Has pastor's name listed

                my ($tmpFirst,
                    $tmpMI,
                    $tmpLast,
                    $tmpSuffix,
                    $tmpSpFirst,
                    $tmpSpMI,
                    $tmpSpLast
                    ) = parseNames( $donorNameFirstLU, $donorNameLastLU );

                $orgContactPers = trim('Pastor '. $tmpFirst.' '.$tmpLast.' '.$tmpSuffix);
            }
            else {
                $orgContactPers = ''; # No pastor listed
            }
            $personType = 'O';
            $donorName = $donorOrgNameLU;
            $donorLastOrg = $donorOrgNameLU;
            $donorFirst = '';
        }
        else { # is individual
            $personType = 'P';
            $donorName = trim($donorNameFirstLU.' '.$donorNameLastLU);

            ( $donorFirst,
              $donorMI,
              $donorLastOrg,
              $donorSuffix,
              $spFirst,
              $spMI,
              $spLast
             ) = parseNames( $donorNameFirstLU, $donorNameLastLU );

            $orgContactPers = '';
        }

        # Append values in this list of variabes to array @tntmpdigrecord
        # This list must be in the same order as the DIG_* constants
        push (@tntmpdigrecord, $acctNoLU, $donorName, $streetLU, $cityLU,
              $stateLU, $zipLU, $contribDateLU, $contribAmtLU, $donorFirst,
              $donorLastOrg, $rcptID_or_YTDAmt, $personType, $orgContactPers,
              $gift_memo, $donorMI, $donorSuffix, $spFirst, $spMI, $spLast,
              $phone, $classLU
              );

        # store our record format
        push(@records, \@tntmpdigrecord);
    }
    # done with the input file
    close($input_fh);


    # Catalogue donations
	#
	# 	Three steps:
	#	1. Eliminate positive/negative pairs of same class
	#	2. Return items to original order
	#	3. Generate unique donation ID and output
	#

	# Phase 1a - record positive and negative offering amounts
	my %positive;
	my %negative;
	{
        foreach my $line (0..$#records) {
    		my $donorID = $records[$line][DIG_DONOR_ID];
    		my $amount = $records[$line][DIG_OFFERING_AMOUNT];
    		my $offeringDate = $records[$line][DIG_OFFERING_DATE];
    		my $class = $records[$line][DIG_CLASS];

    		# skip "n/a" dates ($0.00 amounts)
    		next if($offeringDate eq 'n/a');
            next if($amount eq '');
    		next if($amount eq '0.00');
    		next if($amount == 0);

    		# Positive amounts go into our positive ledger, negative into the other one
    		if($amount > 0) {
    			$amount = $amount * 1;
    			push(@{$positive{$donorID}{$offeringDate}{$amount}{$class}}, $line);
    		}
    		else {
    			$amount = $amount * -1;
    			push(@{$negative{$donorID}{$offeringDate}{$amount}{$class}}, $line);
    		}
        }
	} # End Phase 1a


    # Phase 1b - scan negative offering amounts, and delete those records
	# which have matching positive amounts in the same class, same donor, same date.
	# Delete both the positive and negative pair.
	#
	# AGWM personal donations (class 09) can have a matching negative amount for the same offering date. This appears to happen
	# when a personal donation is given directly to the missionary who is on a field assignment (not reported with itineration reports).
	# In such cases, the negative amount should be ignored, but the positive amount retained. So I will skip deletion of matching
	# positive offerings in class 09.
	{
		# for each donor
		while (my ($donorID, $hashrefA) = each(%negative)) 	{
			# For each donation date
			while (my ($offeringDate, $hashrefB) = each(%$hashrefA)) {
				# For each donation amount
				while (my ($amount, $hashrefC) = each(%$hashrefB)) {
					# For each class
					while (my ($class, $arrayref) = each(%$hashrefC)) {
						my $array_cpy = [@{$arrayref}]; # copy array so that deletion of array elements won't cause premature termination of loop - mrb
						foreach(@$array_cpy) {
							# If a matching positive record exists, pop the donation off both +/- list
							if(exists $positive{$donorID}{$offeringDate}{$amount}{$class}[0]) {
								# Pop off this negative record
								pop @{$negative{$donorID}{$offeringDate}{$amount}{$class}};

                                # Pop off positive record, unless class '09'
								if ($class ne '09') {
									pop @{$positive{$donorID}{$offeringDate}{$amount}{$class}};
								}
							}
						}
					}
				}
			}
		}
	} # End Phase 1b


    # Phase 2 - Return items to original order
    #         - get the ledger returned with unique donations
	@records = get_ledger_with_unique_donations(\@records, \%positive, \%negative);




    # Phase 3 - Prepare output string

    # First declare some variables for storing earliest and latest gift dates
    # They are needed at the time of final output, so are declared outside of
    # the following lexical scope
    my ($earliest_date, $latest_date);

    {
        # A place to keep track of which donors we have written out
        my %donors_written;

        # Loop through each record in our ledger
        foreach my $record (@records) {

            # skip records that we have not marked to keep
            next unless @$record[DIG_KEEP_RECORD];

            # Retrieve necessary values...

            # record number is to differentiate similar gifts that differ only by class number
            my $record_number   = @$record[DIG_RECORD_NUM];
            my $date            = @$record[DIG_OFFERING_DATE];
            my $donorID         = @$record[DIG_DONOR_ID];
            my $amount          = @$record[DIG_OFFERING_AMOUNT];
            my $donorname       = @$record[DIG_DONOR_NAME];
            my $donoraddr       = @$record[DIG_DONOR_ADDR];
            my $donorcity       = @$record[DIG_DONOR_CITY];
            my $donorstate      = @$record[DIG_DONOR_STATE];
            my $donorzip        = @$record[DIG_DONOR_ZIP];
            my $donorfirst      = @$record[DIG_DONOR_FIRSTNAME];
            my $donorlastorg    = @$record[DIG_DONOR_LASTNAME_ORG];
            my $receiptID       = @$record[DIG_RECEIPT_ID];
            my $giftMemoB       = @$record[DIG_GIFT_MEMO];
			my $class			= @$record[DIG_CLASS];


            # Date for running comparison to determine earliest and latests offering dates
            my $this_date_yyyymmdd;

            my ($offeringDate, $day, $month, $year_four_digit, $year_two_digit);
            # Format offering date (read out date parts and convert to TntConnect format)
            #  Update: TntConnect now allows date formats of either:
            #          M/D/YYYY   or  YYYY-MM-DD
            #
            if ($date =~ m/^([0-9]{2})([0-9]{2})-([0-9]{2})-([0-9]{2})$/) {
                $month = $3;
                $day = $4;
                $year_four_digit = $1 . $2;
                $year_two_digit = $2;
                $offeringDate = "$month/$day/$year_four_digit";
                # for date comparisons
                $this_date_yyyymmdd = "$year_four_digit$month$day";
            } else {
                $title = "Warning";
                $dialog = "Internal error: Unexpected date: $date\n" .
                          "Expected date format: yyyy-mm-dd\n\n" .
                          "This error can sometimes be caused if you have opened " .
                          "this CSV file in a spreadsheet for viewing and then " .
                          "saved it when closing the file.  Please try to " .
                          "process an original AGWM CSV file that has not " .
                          "been modified by another application.\n";
                print STDERR wordwrap($dialog);

                if ( is_mac() ) {
                    $dialog = $dialog . "Program will exit now.";
                    dialog_exit($dialog, $title);
                } else {
                    key_exit();
                }
            }


            # Determine date range of gifts by string comparison of dates
            # in YYYYMMDD format
            if (!$earliest_date || $this_date_yyyymmdd lt $earliest_date) {
                $earliest_date = $this_date_yyyymmdd;
            }
            if (!$latest_date || $this_date_yyyymmdd gt $latest_date) {
                $latest_date = $this_date_yyyymmdd;
            }


            # generate donation id
            # Run an MD5 against many fields, to generate "prehash";
            # Since monthly files lack $receiptID, I substituted YTDAmt for this value
            my $prehash = md5_hex("$donorID$offeringDate$amount$receiptID$record_number");

            # A substring of the hash is converted from hex to numeric values
            # It is then prepended by the two-digit year and two-digit month.
            # With using a substring, there is a small possibility of a
            # donation ID collision (duplicate ID). However, it is highly
            # unlikely due to the MD5 avalanche effect.
            #
            # Final string is about 14-15 characters long.
            # TntConnect limits this field to 18 alphanumeric characters.
            $prehash =~ /^(.{8})/;
            my $id = hex($1);
            my $donationid = "$year_two_digit$month-$id";

            # Prefix $donorID with the 'Z_' if it has a leading zero.
            #  This is a hack to prevent TntConnect from deleting the leading zero.
            #  (AGWM donor IDs occasionally have leading zeros, which need to
            #  be preserved to distinguish separate donors)
            if (( $donorID ne '') && (( substr $donorID, 0, 1) eq '0' )) { # begins with 0
                $donorID = 'Z_'. $donorID; # Add prefix Z_
            }

            # Save a long string containing all the gifts for output
                    #   "PEOPLE_ID","ACCT_NAME", "DISPLAY_DATE",
            $gifts .= qq("$donorID","$donorname","$offeringDate",)
                    #    "AMOUNT", "DONATION_ID","DESIGNATION",
                    . qq("$amount","$donationid","1",)
                    #  "MOTIVATION","MEMO"
                    . qq("$class","$giftMemoB"\n);

            # A long string with donor's information
            # If this donor's information hasn't been written to the string yet, then include it
            if(!$donors_written{$donorID}) {
                        #     "PEOPLE_ID","ACCT_NAME", "ADDR1",
                $donors .= qq("$donorID","$donorname","$donoraddr",)
                        # "ADDR2","ADDR3","ADDR4",
                        . qq(,,,)
                        #      "CITY",      "STATE",     "ZIP",
                        . qq("$donorcity","$donorstate","$donorzip",)
                        #     "ADDR_CHANGED","PHONE",               "PHONE_CHANGED",
                        . qq("$offeringDate","@$record[DIG_PHONE]","$offeringDate",)
                        # "COUNTRY","CNTRY_DESCR","PERSON_TYPE",
                        . qq(,,"@$record[DIG_PERSON_TYPE]",)
                        #    "LAST_NAME_ORG", "FIRST_NAME",  "MIDDLE_NAME",
                        . qq("$donorlastorg","$donorfirst","@$record[DIG_DONOR_MI]",)
                        # "TITLE",      "SUFFIX",                "SP_LAST_NAME",
                        . qq(,"@$record[DIG_DONOR_SFFX]","@$record[DIG_SP_LAST]",)
                        #             "SP_FIRST_NAME", "SP_MIDDLE_NAME","SP_TITLE",
                        . qq("@$record[DIG_SP_FIRST]","@$record[DIG_SP_MI]",,)
                        # "SP_SUFFIX", "ORG_CONTACT_PERSON"
                        . qq(,"@$record[DIG_ORG_CONTACT_PERS]"\n);
            }

            # record that this donor's information has been recorded to avoid duplication
            $donors_written{$donorID} = $donorID;

        } # end of this record (A "line" in our ledger)

    } # end of Phase 3


    # Set and check GiftsDateFrom and GiftsDateTo strings
    my ($gifts_date_from,
        $gifts_date_to,
        $gifts_year_from,
        $gifts_year_to,
        $gifts_month_from,
        $gifts_month_to);

    if ($earliest_date =~ m/^([0-9]{4})([0-9]{2})([0-9]{2})$/) {
        $gifts_date_from = "$1-$2-$3";   # YYYY-MM-DD
        $gifts_year_from = $1;
        $gifts_month_from = $2;
    } else {
        print STDERR "Unable to match GiftsDateFrom regex pattern for: " .
            "$earliest_date. Exiting...\n";
            sleep(3);
            die;
    }

    if ($latest_date =~ m/^([0-9]{4})([0-9]{2})([0-9]{2})$/) {
        $gifts_date_to = "$1-$2-$3";   # YYYY-MM-DD
        $gifts_year_to = $1;
        $gifts_month_to = $2;
    } else {
        print STDERR "Unable to match GiftsDateTo regex pattern for: " .
            "$latest_date. Exiting...\n";
            sleep(3);
            die;
    }

    # Now check that gifts date range does not exceed one month
    if ($gifts_year_from ne $gifts_year_to ||
        $gifts_month_from ne $gifts_month_to ) {

            $dialog = "Unexpected Gifts Date Range.\n" .
                "The range of dates for gifts in this file appear to exceed " .
                "one month. If TntConnect has other gifts within this date " .
                "range that are not included in this datasync file, then it " .
                "may prompt you to delete them. Please be aware of this and " .
                "exercise caution.\n" .
                "GiftsDateFrom: $gifts_date_from\n" .
                "GiftsDateTo: $gifts_date_to\n";
			$title = "Alert";
			print STDERR wordwrap($dialog);

			if ( is_mac() ) {
				dialog_ok($dialog, $title);
			} else {
				key_continue();
			}
    } # end date range check


    # Set output filename
    #  1. Remove .csv extension
    #  2. Append '_monthly' if input file was monthly summary
    #  3. Append '.tntdatasync' extension
    my $output_filename = get_filename_without_extension($input_filename);
    if($filetype eq 'monthly') {
        $output_filename .= '_monthly';
    }
    $output_filename .= '.tntdatasync';



    # Processing... message
	print "Outputting file ($output_filename)...\n";

    # Open file for output
	open OUTPUT, ">", "$output_filename" or die "Cannot open file for output. $!";

	my $orgName = 'Assemblies of God World Missions';
	my $orgAbbreviation = 'AGWM';
	my $orgCode = 'AGWM';

	# header
	my $output = <<HEREDOC;
[ORGANIZATION]
Name=$orgName
Abbreviation=$orgAbbreviation
Code=$orgCode
DefaultCurrencyCode=USD

[FILE]
GiftsDateFrom=$gifts_date_from
GiftsDateTo=$gifts_date_to

[DONORS]
"PEOPLE_ID","ACCT_NAME","ADDR1","ADDR2","ADDR3","ADDR4","CITY","STATE","ZIP","ADDR_CHANGED","PHONE","PHONE_CHANGED","COUNTRY","CNTRY_DESCR","PERSON_TYPE","LAST_NAME_ORG","FIRST_NAME","MIDDLE_NAME","TITLE","SUFFIX","SP_LAST_NAME","SP_FIRST_NAME","SP_MIDDLE_NAME","SP_TITLE","SP_SUFFIX","ORG_CONTACT_PERSON"
${donors}
[GIFTS]
"PEOPLE_ID","ACCT_NAME","DISPLAY_DATE","AMOUNT","DONATION_ID","DESIGNATION","MOTIVATION","MEMO"
$gifts
HEREDOC
	chomp($output);
	chomp($output);
	print OUTPUT $output;
	close OUTPUT;

} # End of FILE block. Will continue loop for any additional files passed as arguments
