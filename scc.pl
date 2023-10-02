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
use utf8;
use open qw(:std :utf8);
use Encode;
use Encode::Guess;

#### ADAPTATION ZONE 
my $user = 'eurolinux@gmail.com';
my $timezone = 'Europe/Paris';
# TBC
my $spreadsheet = "1nF3lhXj9U7IrVdu_fEwB7nKTq057D2iFPDA9_q5kbME1nF3lhXj9U7IrVdu_fEwB7nKTq057D2iFPDA9_q5kbME";
# Start in check mode, and pass it to 0 once verified
my $check = 0;
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
#use Google::RestApi;

#print "Step 1\n";
#my $rest_api = Google::RestApi->new( config_file   => "config.yaml");
# content seems OK
#print "RA: ".Dumper($rest_api)."\n";
 
#print "Step 2\n";
#use Google::RestApi::SheetsApi4;
#print "Step 3\n";
#my $sheets_api = Google::RestApi::SheetsApi4->new(api => $rest_api);
#print "SA: ".Dumper($sheets_api)."\n";
#print "Step 4\n";
#my $sheet = $sheets_api->open_spreadsheet(title => "Sélection SCC 22/23");
##/edit#gid=968175567
#my $sheet = $sheets_api->open_spreadsheet(id => $spreadsheet);
#print "SH: ".Dumper($sheet)."\n";
#print "Step 5\n";
# Doesn't work
#my $ws0 = $sheet->open_worksheet(id => "968175567");
#my $ws0 = $sheet->open_worksheet(name => "2022-2023");
#print "WS ".Dumper($ws0)."\n";
## doesn't work
#my $ws0 = $sheet->open_worksheet(id => 1);
#print "Found WS0 ".Dumper($ws0)."\n";

#print "Step 6\n";
#my $rc = "E2";
#my $cell = $ws0->range_cell($rc);
#print "CELL ".Dumper($ws0)."\n";
#print "Step 6-2\n";
# Doesn't work
#my $ws0id = $cell->worksheet_id();
#print "WS ID : $ws0id\n";
#print "Step 6-3\n";
#my $ssid = $cell->spreadsheet_id();
#print "WS ID : $ssid\n";
#print "Step 6-4\n";
#print "Found $rc ".Dumper($cell)."\n";
#print "Step 6-5\n";
#print "Values:\n".Dumper($cell->values())."\n";

#print "Step 6-6\n";
#my $vals = $ws0->rows([1, 2, 3]);
#print "Vals ".Dumper($vals)."\n";

#print "Step 7\n";
#my $cols = $ws0->tie_cols('SPECTACLE', 'PRECISION', 'SALLES');
#print "Cols: $cols->{SPECTACLE} -> $cols->{SALLES}\n";

#print "Step 8\n";
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
my $calendar_id;
my $calendar_only_id;
#my $calendar_name;

#
# This GOOGLE API works for calendar !
#
print "Step 0\n";
# Use goauth for that
my $gapi = API::Google::GCal->new({ tokensfile => 'config.json' });

print "Step 1\n";
$gapi->refresh_access_token_silent($user); # inherits from API::Google

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
		# Spectacles Auto
		$calendar_id = $gapi->get_calendar_id_by_name($user, 'SpectaclesAuto');
		# TEST
		#$calendar_id = $gapi->get_calendar_id_by_name($user, 'TEST');
		last;
	} 
	if ($t->get_name =~ /202.*/)  {
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
		$fields{'cat'} = "A";
		$fields{'style'} = "B";
		$fields{'options'} = "U";
		$fields{'url'} = "AE";
		$fields{'text'} = "AF";
		$fields{'choix'} = "K";
		# TEST
		#$calendar_id = $gapi->get_calendar_id_by_name($user, 'TEST');
		#$calendar_only_id = $calendar_id;
		# SCC
		$calendar_id = $gapi->get_calendar_id_by_name($user, 'SCC');
		# Coups de Cœur
		$calendar_only_id = $gapi->get_calendar_id_by_name($user, 'Coups de Cœur');
		last;
	}
}
print "Found spreadhseet of type $scctype\n" if (defined $scctype);
if (defined $calendar_id) {
	print "Will populate calendar $calendar_id\n";
} else {
	print "Use goauth to re-enable Google Calendar connection\n";
    exit(-1);
}

# Capture all scc
print "Analyzing it ...\n" if (defined $scctype);
my $scc = $odf->get_body->get_table($worksheet);

my $i = 1;
my $c;
my $end = 0;
my $cell;
while ($end eq 0) {
	foreach my $k (keys %fields) {
		$cell = $scc->get_cell($i, $fields{$k});
		$end = 1 if ((not defined $cell) and ($k =~ /spectacle/));
		$c = $cell->get_text();
		$end = 1 if ((not defined $c) || ($c eq "") and ($k =~ /spectacle/));
		$cal{$i}->{$k} = $cell->get_text();
	}
	# Manages lack of minutes
	$cal{$i}{'duration'} .= "0" if ((defined $cal{$i}{'duration'}) and ($cal{$i}{'duration'} =~ /h$/));
	# Skipping Choices 2+ for SCC
	delete($cal{$i}) if (($scctype eq "ICE") and ($cal{$i}{'choix'} !~ /1/));
	$i++;
}
# The previous one is void delete it
$i--;
delete($cal{$i});
#print "Calendar :\n";
#print Dumper(%cal);
print "... Done.\n" if (defined $scctype);

my ($dateparser,$dateparserend,$duration,$event_start,$event_end);

my $j = 0;
foreach my $i (sort keys %cal) {
	#print Dumper($cal{$i});
	$event->{description} = "$cal{$i}->{spectacle}\n\n$cal{$i}->{detail}\n";
	$event->{description} .= "Catégorie: $cal{$i}->{cat}\n" if (defined $cal{$i}->{cat});
	$event->{description} .= "Style $cal{$i}->{style}\n" if (defined $cal{$i}->{style});
	$event->{description} .= "Détails: $cal{$i}->{text}\n" if (defined $cal{$i}->{text});
	$event->{description} .= "URL $cal{$i}->{url}\n" if (defined $cal{$i}->{url});
	$event->{description} .= "Nb options $cal{$i}->{options}\n" if (defined $cal{$i}->{options});
	$event->{description} = decode("Guess", $event->{description});
	$event->{summary} = decode("Guess","$cal{$i}->{spectacle}");
	$event->{location} = decode("Guess","$cal{$i}->{salle}");
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
	# This is for the event itself
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
	print "Add event #$j $cal{$i}->{spectacle} - $event->{start}{dateTime}\n";
	#print Dumper($event);
	$gapi->add_event($user, $calendar_id, $event) if ($check eq 0);
	# For SCC we also need reminders for options, sending confirmation, and invoice
	if ($scctype eq "ICE") {
		$gapi->add_event($user, $calendar_only_id, $event) if ($check eq 0);
		# First manages options at the planned date
		$event->{summary} = decode("Guess","Rendu d'options pour $cal{$i}->{spectacle}");
		$event_start = $dateparser->parse_datetime($cal{$i}->{dateinsc}." 18h00");
		$event->{start}{timeZone} = $timezone;
		$event->{start}{dateTime} = DateTime::Format::RFC3339->format_datetime($event_start);
		$event_end = $event_start + DateTime::Duration->new( hours => 1 );
		$event->{end}{dateTime} = DateTime::Format::RFC3339->format_datetime($event_end);
		print "Add reminder option event #$j $cal{$i}->{spectacle}\n";
		$gapi->add_event($user, $calendar_id, $event) if ($check eq 0);

		# Then manages participant communication one week before date
		$event->{summary} = decode("Guess","Participants pour $cal{$i}->{spectacle}");
		$event_start = $dateparser->parse_datetime($cal{$i}->{date}." 18h00");
		$event_start = $event_start + DateTime::Duration->new( weeks => -1 );
		$event->{start}{timeZone} = $timezone;
		$event->{start}{dateTime} = DateTime::Format::RFC3339->format_datetime($event_start);
		$event_end = $event_start + DateTime::Duration->new( hours => 1 );
		$event->{end}{dateTime} = DateTime::Format::RFC3339->format_datetime($event_end);
		print "Add participant option event #$j $cal{$i}->{spectacle}\n";
		$gapi->add_event($user, $calendar_id, $event) if ($check eq 0);
		
		# Then manages participants invoice one week after
		$event->{summary} = decode("Guess","Facture participants pour $cal{$i}->{spectacle}");
		$event_start = $dateparser->parse_datetime($cal{$i}->{date}." 18h00");
		$event_start = $event_start + DateTime::Duration->new( weeks => 1 );
		$event->{start}{timeZone} = $timezone;
		$event->{start}{dateTime} = DateTime::Format::RFC3339->format_datetime($event_start);
		$event_end = $event_start + DateTime::Duration->new( hours => 1 );
		$event->{end}{dateTime} = DateTime::Format::RFC3339->format_datetime($event_end);
		print "Add participant invoice option event #$j $cal{$i}->{spectacle}\n";
		$gapi->add_event($user, $calendar_id, $event) if ($check eq 0);
	}

}
