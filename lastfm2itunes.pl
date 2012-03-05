#!/usr/bin/perl

use strict;
use XML::DOM;
use Class::Struct;
use Encode;
use POSIX qw(strftime);

use Win32::OLE 'CP_UTF8';
$Win32::OLE::CP = CP_UTF8;
my $parser = new XML::DOM::Parser;

# User static settings
my $username = 'TedIrens';		# <= Put your username here
my $verbose = 0;		# <= set 1 to enable verbose output

# Script static variables
my $chart_list_url   = 'http://ws.audioscrobbler.com/2.0/user/<USER>/weeklychartlist.xml';
my $weekly_chart_url = 'http://ws.audioscrobbler.com/2.0/user/<USER>/weeklytrackchart.xml?from=<FROM>&to=<TO>';
my $version = 'v0.2 (05.03.2012)';

# Script dynamic variables
my %lastfm_track_playcount = ();
my %lastfm_track_playlast = ();

print "iTunes Library Updater " . $version . "\n";
print "Copyright (c) 2012 Roman Gemini (roman_gemini\@ukr.net)\n\n";
if($username eq '') {
	print "Please, set variable '\$username' first!\n";
	print "Press <ENTER> to exit...";
	<>;
	exit;
}

print "Trying to connect to iTunes scripting interface...\n";
my $iTunes = Win32::OLE->GetActiveObject('iTunes.Application');
unless ($iTunes) { $iTunes = new Win32::OLE('iTunes.Application'); }
unless ($iTunes) { die("Couldn't connect to iTunes!"); }

print "Fetching week charts...\n";
my $url = $chart_list_url;
$url =~ s/<USER>/$username/g;
my $charts_data = $parser->parsefile($url) || die("Can't fetch charts list XML!");
my $root_charts_data = $charts_data->getElementsByTagName("weeklychartlist") || die("Can't parse charts list XML!");
my $weeks_charts_data = $root_charts_data->item(0)->getElementsByTagName("chart") || die("Can't parse charts list XML!");
for my $i (0 .. $weeks_charts_data->getLength-1) {
	my $item = $weeks_charts_data->item($i);
	my $from = $item->getAttribute("from");
	my $to = $item->getAttribute("to");

	my $w_url = $weekly_chart_url;
	$w_url =~ s/<USER>/$username/g;
	$w_url =~ s/<FROM>/$from/g;
	$w_url =~ s/<TO>/$to/g;
	print sprintf("Fetching chart for week %3d of %3d (%s - %s)\n", $i + 1, $weeks_charts_data->getLength, strftime("%Y-%m-%d", localtime($from)), strftime("%Y-%m-%d", localtime($to)));

	my $week_data;
	eval { $week_data = $parser->parsefile($w_url) };
	next if(!$week_data); 

	my $week_root = $week_data->getElementsByTagName("weeklytrackchart");
	my $week_tracks = $week_root->item(0)->getElementsByTagName("track");

	for my $j (0 .. $week_tracks->getLength-1) {
		my $track = $week_tracks->item($j);
		my $tag_title = $track->getElementsByTagName("name")->item(0)->getFirstChild->getData();
		my $tag_artist = $track->getElementsByTagName("artist")->item(0)->getFirstChild->getData();
		my $tag_play_count = $track->getElementsByTagName("playcount")->item(0)->getFirstChild->getData();
		print sprintf("[v] Track \"%s\" played %d time(s)\n", _866("$tag_artist - $tag_title"), $tag_play_count) if($verbose);

		$lastfm_track_playcount{lc($tag_artist)}{lc($tag_title)} += $tag_play_count;
		$lastfm_track_playlast{lc($tag_artist)}{lc($tag_title)} = strftime("%Y-%m-%d %H:%M:%S", localtime($to));
	}
}

print "\nProcessing iTunes library...\n";
my $iTunes_LIB = $iTunes->LibraryPlaylist->Tracks;
if($iTunes_LIB) {
	for my $t (1 .. $iTunes_LIB->Count) {
		my $trk = $iTunes_LIB->Item($t);
		next unless($trk->kind() == 1);

		next unless(exists($lastfm_track_playcount{lc($trk->artist())}{lc($trk->name())}));

		my $tmp_count = $lastfm_track_playcount{lc($trk->artist())}{lc($trk->name())};
		my $tmp_last = $lastfm_track_playlast{lc($trk->artist())}{lc($trk->name())};

		if($trk->playedCount() < $tmp_count) {
			printf("Updating track %5d: \"%s\"\n", $trk->trackID(), _866($trk->artist() . " - " . $trk->name()));
			$trk->{playedCount} = $tmp_count;
			$trk->{playedDate} = $tmp_last;
		} else {
			printf("Skipping track %5d: \"%s\"\n", $trk->trackID(), _866($trk->artist() . " - " . $trk->name())) if($verbose);
		}
	}
}

undef $charts_data;
undef $parser;
undef $iTunes;
print "Done! Press <ENTER> to continue...\n";
<>;

sub _866 { return encode("cp866", shift); }