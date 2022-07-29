#!/usr/bin/perl -w
#
use strict;
use Data::Dumper;
use API::Google::GCal;

#### ADAPTATION ZONE 
my $user = 'eurolinux@gmail.com';
# TEST
#my $calendar_id = '08ulfn9rcjdl3c5spj12n6cvk4@group.calendar.google.com';
# SCC
my $calendar_id = '228v2ibpe7hggi0rdb13n2drk8@group.calendar.google.com';
# Coup de Cœur only
#my $calendar_id = '069oro5hmku3vou5hjafdj66a0@group.calendar.google.com';
# Start in check mode, and pass it to 0 once verified
my $check = 0;
#### END ZONE

# Use goauth for that
my $gapi = API::Google::GCal->new({ tokensfile => 'config.json' });

#print Dumper($gapi);
#print "Step 1\n";
$gapi->refresh_access_token_silent($user); # inherits from API::Google
my $c = $gapi->get_calendar_name($user, $calendar_id);
print "Purging all events from calendar $c ...\n";
print "Waiting 5 seconds in case it's not what you want !\n";
sleep(5);

#print "Step 2\n";
my $e = $gapi->events_list({calendarId => $calendar_id, user => $user});
for my $id (@$e) {
	#print Dumper($id);
	print "Deleting event $id->{'summary'}\n";
	$gapi->delete_event($user, $calendar_id, $id->{'id'}) if ($check eq 0);
}
