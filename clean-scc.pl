#!/usr/bin/perl -w
#
use strict;
use Data::Dumper;
use API::Google::GCal;

#### ADAPTATION ZONE 
my $user = 'eurolinux@gmail.com';
# TBC
my $calendar_id = '08ulfn9rcjdl3c5spj12n6cvk4@group.calendar.google.com';
# Start in check mode, and pass it to 0 once verified
my $check = 0;
#### END ZONE

print "Purging all events from calendar ...\n";
# Use goauth for that
my $gapi = API::Google::GCal->new({ tokensfile => 'config.json' });

print "Step 1\n";
$gapi->refresh_access_token_silent($user); # inherits from API::Google

print "Step 2\n";
my $e = $gapi->events_list({calendarId => $calendar_id, user => $user});
for my $id (@$e) {
	#print Dumper($id);
	print "Deleting event $id->{'summary'}\n";
	$gapi->delete_event($user, $calendar_id, $id->{'id'}) if ($check eq 0);
}
