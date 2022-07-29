#!/usr/bin/perl -w
#
use strict;
use ODF::lpOD;
use Data::Dumper;
use DateTime;
use DateTime::Format::RFC3339;
use DateTime::Format::Strptime;
use DateTime::Format::Duration;
use API::Google::GCal;

#### ADAPTATION ZONE 
my $user = 'eurolinux@gmail.com';
my $calendar_id = '08ulfn9rcjdl3c5spj12n6cvk4@group.calendar.google.com';
my $calendar_name = "TEST";
my $timezone = 'Europe/Paris';
# TBC
my $spreadsheet = "1nF3lhXj9U7IrVdu_fEwB7nKTq057D2iFPDA9_q5kbME1nF3lhXj9U7IrVdu_fEwB7nKTq057D2iFPDA9_q5kbME";
#### END ZONE

#Avoids warnings
{
    package ODF::lpOD::Element;
    no warnings 'once';
    *DESTROY = sub {} unless defined &DESTROY;
}

STDOUT->autoflush(1);


### GOOGLE Read NOT WORKING for now ###
# Use google_restapi_session_creator to create token with config.yaml from Google console
use Google::RestApi;

print "Step 1\n";
my $rest_api = Google::RestApi->new( config_file   => "config.yaml");
# content seems OK
#print "RA: ".Dumper($rest_api)."\n";
 
print "Step 2\n";
use Google::RestApi::SheetsApi4;
print "Step 3\n";
my $sheets_api = Google::RestApi::SheetsApi4->new(api => $rest_api);
print "SA: ".Dumper($sheets_api)."\n";
print "Step 4\n";
#my $sheet = $sheets_api->open_spreadsheet(title => "Sélection SCC 22/23");
##/edit#gid=968175567
my $sheet = $sheets_api->open_spreadsheet(id => $spreadsheet);
print "SH: ".Dumper($sheet)."\n";
print "Step 5\n";
# Doesn't work
#my $ws0 = $sheet->open_worksheet(id => "968175567");
my $ws0 = $sheet->open_worksheet(name => "2022-2023");
print "WS ".Dumper($ws0)."\n";
## doesn't work
#my $ws0 = $sheet->open_worksheet(id => 1);
#print "Found WS0 ".Dumper($ws0)."\n";

print "Step 6\n";
my $rc = "E2";
my $cell = $ws0->range_cell($rc);
print "CELL ".Dumper($ws0)."\n";
print "Step 6-2\n";
# Doesn't work
#my $ws0id = $cell->worksheet_id();
#print "WS ID : $ws0id\n";
print "Step 6-3\n";
my $ssid = $cell->spreadsheet_id();
print "WS ID : $ssid\n";
print "Step 6-4\n";
print "Found $rc ".Dumper($cell)."\n";
print "Step 6-5\n";
#print "Values:\n".Dumper($cell->values())."\n";

print "Step 6-6\n";
#my $vals = $ws0->rows([1, 2, 3]);
#print "Vals ".Dumper($vals)."\n";

print "Step 7\n";
#my $cols = $ws0->tie_cols('SPECTACLE', 'PRECISION', 'SALLES');
#print "Cols: $cols->{SPECTACLE} -> $cols->{SALLES}\n";

print "Step 8\n";
#my $values = $ws0->rows([2, 3, 4]);
#foreach my $v ($values) {
#print Dumper($v);
#}
#
### END GOOGLE Read NOT WORKING for now ###


my %cal; # Calendar data structure
my $event = {}; # Event data structure

#
# Play with the ods file downloaded instead
#
print "\nStarting with SCC data collection\n\n";
# Input management
if (not defined $ARGV[0]) {
        print "Syntax: scc.pl file-scc.ods\n";
        exit(-1);
}
open(CSV,"$ARGV[0]") || die "Unable to read $ARGV[0]";
close(CSV);
# ods management
my $odf = odf_document->get($ARGV[0]);
if (not defined $odf) {
        print "File $ARGV[0] isn't defined\n";
        exit(-1);
}

my $context = $odf->body;
my @tables = $context->get_tables;
my $scctype;
my $worksheet;
my $salles;
my %fields;

# Detect SSC type : ICE or Perso
foreach my $t (@tables) {
	if ($t->get_name =~ /Spectacles/) {
		# Perso
		$scctype = "Perso";
		$worksheet = $t->get_name;
		$salles = "Salles";
		# Map colunms
		$fields{'spectacle'} = "A";
		$fields{'detail'} = "B";
		$fields{'duration'} = "C";
		$fields{'start'} = "G";
		$fields{'date'} = "H";
		$fields{'salle'} = "J";
	} elsif ($t->get_name =~ /2022-2023/)  {
		# ICE
		$scctype = "ICE";
		$worksheet = $t->get_name;
		$salles = "Relais de Salles";
		# Map colunms
		$fields{'spectacle'} = "E";
		$fields{'detail'} = "F";
		$fields{'duration'} = "J";
		$fields{'start'} = "Q";
		$fields{'date'} = "R";
		$fields{'salle'} = "G";
		$fields{'dateinsc'} = "S";
	}
}
print "Found spreadhseet of type $scctype\n" if (defined $scctype);

# Capture all scc
my $scc = $odf->get_body->get_table($worksheet);

my $i = 1;
my $c;
my $end = 0;
while ($end eq 0) {
	foreach my $k (keys %fields) {
		$cell = $scc->get_cell($i, $fields{$k});
		$end = 1 if ((not defined $cell) and ($k =~ /spectacle/));
		$c = $cell->get_text();
		$end = 1 if ((not defined $c) || ($c eq "") and ($k =~ /spectacle/));
		$cal{$i}->{$k} = $cell->get_text();
	}
	# Manages lack of minuts
	$cal{$i}{'duration'} .= "0" if ((defined $cal{$i}{'duration'}) and ($cal{$i}{'duration'} =~ /h$/));
	$i++;
}
# The previous one is void delete it
$i--;
delete($cal{$i});
#print "Calendar :\n";
#print Dumper(%cal);

#
# This GOOGLE API works for calendar !
#
print "Step 0\n";
my $gapi = API::Google::GCal->new({ tokensfile => 'config.json' });


print "Step 1\n";
$gapi->refresh_access_token_silent($user); # inherits from API::Google

#print "Step 2\n";
#$gapi->get_calendars($user);
#print "Step 3\n";
#$gapi->get_calendars($user, ['id', 'summary']);  # return only specified fields

print "Step 2\n";
$gapi->get_calendar_id_by_name($user, $calendar_name);

my ($dateparser,$dateparserend,$duration,$event_start,$event_end);

my $j = 0;
foreach my $i (sort keys %cal) {
	print Dumper($cal{$i});
	$event->{description} = "$cal{$i}->{spectacle}\n$cal{$i}->{detail}\n";
	$event->{summary} = "$cal{$i}->{spectacle}";
	$event->{location} = "$cal{$i}->{salle}";
	if ($scctype eq "Perso") {
		$dateparser = DateTime::Format::Strptime->new( 
			pattern => '%d/%m/%Y %Hh%M',
			time_zone => $timezone,
			on_error => 'croak',
		);
	}
	if ($scctype eq "ICE") {
		$dateparser = DateTime::Format::Strptime->new( 
			pattern => '%Y-%m-%d %Hh%M',
			time_zone => $timezone,
			on_error => 'croak',
		);
	}
	$event_start = $dateparser->parse_datetime($cal{$i}->{date}." ".$cal{$i}->{start});
	$event_end = $dateparser->parse_datetime($cal{$i}->{date}." ".$cal{$i}->{start});
	$event->{start}{timeZone} = $timezone;
	# '2016-11-11T09:00:00+03:00' format
	$event->{start}{dateTime} = DateTime::Format::RFC3339->format_datetime($event_start);
	$event->{end}{timeZone} = $timezone;
	$dateparserend = DateTime::Format::Duration->new( 
			pattern => '%Hh%M',
		);
	$duration = $dateparserend->parse_duration($cal{$i}->{duration});
	$event_end += $duration;
	$event->{end}{dateTime} = DateTime::Format::RFC3339->format_datetime($event_end);

	$j++;
	print "Add event #$j $cal{$i}->{spectacle}\n";
	print Dumper($event);
	$gapi->add_event($user, $calendar_id, $event);
}
