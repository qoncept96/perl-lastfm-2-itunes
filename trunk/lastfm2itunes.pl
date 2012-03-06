#!/usr/bin/perl

use FindBin qw($Bin);
use strict;
use XML::DOM;
use Class::Struct;
use Encode;
use POSIX qw(strftime);
use LWP::Simple;
use Time::Local;
use Time::HiRes qw(usleep);

use Win32::OLE::Variant;
use Win32::OLE 'CP_UTF8';
$Win32::OLE::CP = CP_UTF8;
my $parser = new XML::DOM::Parser;

# User static settings
my $username = 'TedIrens';		# <= Put your username here
my $verbose = 0;		# <= set 1 to enable verbose output
my $use_cache = 1;		# <= set 1 to cache Last.fm data

# Script static variables
my $chart_list_url    = 'http://ws.audioscrobbler.com/2.0/user/<USER>/weeklychartlist.xml';
my $weekly_chart_url  = 'http://ws.audioscrobbler.com/2.0/user/<USER>/weeklytrackchart.xml?from=<FROM>&to=<TO>';
my $cache_file_format = '\\<USER>-<ID>-<FROM>-<TO>.xml';
my $version = 'v0.3 (06.03.2012)';

# Script dynamic variables
my %lastfm_track_playcount = ();
my %lastfm_track_playlast = ();

# statistic variables
my $processed_tracks = 0;
my $skipped_tracks = 0;

# Cache section
my $cache_dir = $Bin . "\\cache";
if($use_cache) {
	mkdir $cache_dir unless(-d $cache_dir);
}

print "iTunes Library Updater " . $version . "\n";
print "Copyright (c) 2012 Roman Gemini (roman_gemini\@ukr.net)\n\n";
if($username eq '') {
	print "Please, set variable '\$username' first!\n";
	print "Press <ENTER> to exit...";
	<>;
	exit;
}
print "Debug messages : " . ($verbose ? "YES" : "NO") . "\n";
print "Use cache      : " . ($use_cache ? "YES" : "NO") . "\n\n";

print "Trying to connect to iTunes scripting interface...";
my $iTunes = Win32::OLE->GetActiveObject('iTunes.Application');
unless ($iTunes) { $iTunes = new Win32::OLE('iTunes.Application'); }
unless ($iTunes) { print "Error!\n"; die("Couldn't connect to iTunes!"); }
print "OK\n";

print "\nFetching week charts...\n";
my $url = $chart_list_url;
$url =~ s/<USER>/$username/g;
my $charts_data = $parser->parsefile($url);
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

	my $cache_file = $cache_dir . $cache_file_format;
	$cache_file =~ s/<USER>/$username/g;
	$cache_file =~ s/<FROM>/$from/g;
	$cache_file =~ s/<TO>/$to/g;
	$cache_file =~ s/<ID>/($i+1)/g;

	print sprintf("Reading weekly track chart for week %3d of %3d...", $i + 1, $weeks_charts_data->getLength);

	my $raw_data; 
	my $cache_trigr = 0;

	if(-e $cache_file) {
		open XML, "<", $cache_file;
		$raw_data = join('', <XML>);
		close XML;
	} else {
		unless($raw_data = get($w_url)) {
			print "Error!\n";
			die("Can't download xml data!");
		}
		$cache_trigr = 1;
		usleep(250);
	}

	my $week_data;
	eval { $week_data = $parser->parse($raw_data) };
	unless($week_data) {
		print "Invalid XML data!\n";
		next;
	}
	
	print "OK\n";

	# Save cache file
	if(($i < $weeks_charts_data->getLength-1) and $use_cache and $cache_trigr) {
		open XML, ">", $cache_file;
		binmode XML, ":utf8";
		print XML $raw_data;
		close XML;
	}

	my $week_root = $week_data->getElementsByTagName("weeklytrackchart");
	my $week_tracks = $week_root->item(0)->getElementsByTagName("track");

	printf("XML file size: %d bytes, number of tracks: %d\n", length($raw_data), $week_tracks->getLength) if($verbose);

	for my $j (0 .. $week_tracks->getLength-1) {
		my $track = $week_tracks->item($j);
		my $tag_title = $track->getElementsByTagName("name")->item(0)->getFirstChild->getData();
		my $tag_artist = $track->getElementsByTagName("artist")->item(0)->getFirstChild->getData();
		my $tag_play_count = $track->getElementsByTagName("playcount")->item(0)->getFirstChild->getData();
		print sprintf("[v] Track \"%s\" played %d time(s)\n", _866("$tag_artist - $tag_title"), $tag_play_count) if($verbose);

		$lastfm_track_playcount{lc($tag_artist)}{lc($tag_title)} += $tag_play_count;
		$lastfm_track_playlast{lc($tag_artist)}{lc($tag_title)} = strftime("%d.%m.%Y %H:%M:%S", localtime($to));
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
		my $processed = 0;

		if($trk->playedCount() < $tmp_count) {
			printf("Updating \"%s\": setting playedCount[%d] => playedCount[%d]\n", _866($trk->artist() . " - " . $trk->name()), $trk->playedCount(), $tmp_count);
			$trk->{playedCount} = $tmp_count;
			$processed = 1;
		}

		if(strtotime($trk->playedDate()) <= strtotime($tmp_last)) {
			printf("Updating \"%s\": setting playedDate[%s] => playedDate[%s]\n", _866($trk->artist() . " - " . $trk->name()), $trk->playedDate(), $tmp_last);
			$trk->{playedDate} = $tmp_last;
			$processed = 1;
		}

		if($processed) {
			$processed_tracks ++;
		} else {
			$skipped_tracks ++;
		}
	}
}

my $l_artists = 0;
my $l_tracks = 0;
my $l_plays = 0;
foreach my $k_artist (keys %lastfm_track_playcount) {
	$l_artists ++;
	for my $k_track (keys %{$lastfm_track_playcount{$k_artist}}) {
		$l_tracks ++;
		$l_plays += $lastfm_track_playcount{$k_artist}{$k_track};
	}
}

print  "+---------- iTunes Library ---------+\n";
printf "| Number of tracks total | %8d |\n", $iTunes_LIB->Count;
printf "| Number of processed    | %8d |\n", $processed_tracks;
printf "| Number of skipped      | %8d |\n", $skipped_tracks;
print  "+---------- Last.fm Charts ---------+\n";
printf "| Number of artists      | %8d |\n", $l_artists;
printf "| Number of tracks       | %8d |\n", $l_tracks;
printf "| Number of scrobbles    | %8d |\n", $l_plays;
print  "+-----------------------------------+\n\n";

print "\nDone! Press <ENTER> to exit...\n";
<>;


sub _866 { 
	return encode("cp866", shift);
}

sub strtotime {
	my $date = shift;
	my @d;
	# YYYY-MM-DD HH:MM:SS
	if(@d = $date =~ m/(\d{4})[-](\d{2})[-](\d{2})\s(\d{2})[:](\d{2})[:](\d{2})/) {	return timelocal(@d[5,4,3,2,1,0]); }
	# DD.MM.YYYY HH:MM:SS
	if(@d = $date =~ m/(\d{2})[\.](\d{2})[\.](\d{4})\s(\d{2})[:](\d{2})[:](\d{2})/) { return timelocal(@d[5,4,3,0,1,2]); }
	# YYYY-MM-DD
	if(@d = $date =~ m/(\d{4})[-](\d{2})[-](\d{2})/) { return timelocal(0, 0, 0, @d[2,1,0]); }
	return -1;
}