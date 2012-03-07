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
my $cache_file = $Bin . "\\cache.dat";

# User static settings
my $username = 'TedIrens';		# <= Put your username here
my $verbose = 0;		# <= set 1 to enable verbose output

# Script static variables
my $recent_tracks_url = 'http://ws.audioscrobbler.com/2.0/user/<USER>/recenttracks.xml?limit=200&page=<PAGE>';
my $version = 'v0.6 (07.03.2012)';

# Script dynamic variables
my %lastfm_track_playcount = ();
my %lastfm_track_playlast = ();

# statistic variables
my $processed_tracks = 0;
my $skipped_tracks = 0;

print "iTunes Library Updater " . $version . "\n";
print "Copyright (c) 2012 Roman Gemini (roman_gemini\@ukr.net)\n\n";
if($username eq '') {
	print "Please, set variable '\$username' first!\n";
	print "Press <ENTER> to exit...";
	<>;
	exit;
}
print "Debug messages : " . ($verbose ? "YES" : "NO") . "\n\n";

print "Trying to connect to iTunes scripting interface...";
my $iTunes = Win32::OLE->GetActiveObject('iTunes.Application');
unless ($iTunes) { $iTunes = new Win32::OLE('iTunes.Application'); }
unless ($iTunes) { print "Error!\n"; die("Couldn't connect to iTunes!"); }
print "\n";

my $page = 1;
my $pages = 1;
my $scrobbles = 0;

# define variables
my $url;
my $total = 0;

my $rc_data;
my $rc_root;
my $rc_tracks;

my $track;
my $tag_title;
my $tag_artist;
my $tag_play_date;
my $last_date = 0;

my $cache_last_date = 0;
load_cache();

print "Processing XML data. This will take a while...\n";
P1: while(1) {

	$url = $recent_tracks_url;
	$url =~ s/<USER>/$username/g;
	$url =~ s/<PAGE>/$page/g;

	eval { $rc_data = $parser->parsefile($url) };
	unless($rc_data) {
		print "Page $page contains errors!\n";
		if($page < $pages) { next; } else { last; }
	}

	$rc_root = $rc_data->getElementsByTagName("recenttracks");

	$pages     = $rc_root->item(0)->getAttribute("totalPages");
	$total     = $rc_root->item(0)->getAttribute("total");
	$rc_tracks = $rc_root->item(0)->getElementsByTagName("track");

	print "Parsing page $page of $pages\n";

	TRK: for my $j (0 .. $rc_tracks->getLength-1) {
		$track = $rc_tracks->item($j);
		next if ($track->getAttribute("nowplaying") eq "true");

		$tag_title = lc($track->getElementsByTagName("name")->item(0)->getFirstChild->getData());
		$tag_artist = lc($track->getElementsByTagName("artist")->item(0)->getFirstChild->getData());
		$tag_play_date = $track->getElementsByTagName("date")->item(0)->getAttribute("uts");
		
	        $last_date = $tag_play_date if($tag_play_date > $last_date);
		last P1 if($tag_play_date <= $cache_last_date);

		if($lastfm_track_playlast{$tag_artist}{$tag_title} < $tag_play_date) {
			$lastfm_track_playlast{$tag_artist}{$tag_title} = $tag_play_date; 
		}

		$lastfm_track_playcount{$tag_artist}{$tag_title} ++;

	}

	if($page >= $pages) {
		last;
	} else {
		$page ++;
	}

}

dump_cache();

print "Processing iTunes library...\n";
my $iTunes_LIB = $iTunes->LibraryPlaylist->Tracks;

my $trk;
my $tmp_count;
my $tmp_last;
my $processed = 0;
my $tmp_utc;
my $trk_artist;
my $trk_title;

if($iTunes_LIB) {
	for my $t (1 .. $iTunes_LIB->Count) {
		$trk = $iTunes_LIB->Item($t);
		$trk_artist = lc($trk->artist());
		$trk_title = lc($trk->name());

		next unless($trk->kind() == 1);
		next unless(exists($lastfm_track_playcount{$trk_artist}{$trk_title}));

		$tmp_count = $lastfm_track_playcount{$trk_artist}{$trk_title};
		$tmp_utc   = $lastfm_track_playlast{$trk_artist}{$trk_title};
		$tmp_last  = timetostr($tmp_utc);
		$processed = 0;

		if($trk->playedCount() != $tmp_count) {
			printf("Updating \"%s\": updating playedCount[%d] => playedCount[%d]\n", _866($trk->artist() . " - " . $trk->name()), $trk->playedCount(), $tmp_count);
			$trk->{playedCount} = $tmp_count;
			$processed = 1;
		}

		if(strtotime($trk->playedDate()) < strtotime($tmp_last)) {
			printf("Updating \"%s\": setting playedDate[%s] => playedDate[%s]\n", _866($trk->artist() . " - " . $trk->name()), $trk->playedDate(), $tmp_last);
			$trk->{playedDate} = timetostrGM($tmp_utc);
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

print  "Done!\n\nStatistics:\n";
print  "+--------------------- iTunes Library ---------------------+\n";
printf "| Number of tracks total                        | %8d |\n", $iTunes_LIB->Count;
printf "| Number of processed                           | %8d |\n", $processed_tracks;
printf "| Number of skipped                             | %8d |\n", $skipped_tracks;
print  "+--------------------- Last.fm Charts ---------------------+\n";
printf "| Number of artists                             | %8d |\n", $l_artists;
printf "| Number of tracks                              | %8d |\n", $l_tracks;
printf "| Number of scrobbles                           | %8d |\n", $l_plays;
print  "+----------------------------------------------------------+\n\n";

undef $iTunes_LIB, $iTunes;

print "Press <ENTER> to exit...\n";
<>;


sub _866 { return encode("cp866", shift); }
sub timetostr {	return strftime("%d.%m.%Y %H:%M:%S", localtime(shift)); }
sub timetostrGM { return strftime("%d.%m.%Y %H:%M:%S", gmtime(shift)); }
sub strtotime {
	my $date = shift;
	my @d;
	# YYYY-MM-DD HH:MM:SS
	if(@d = $date =~ m/(\d{4})[-](\d{2})[-](\d{2})\s(\d{2})[:](\d{2})[:](\d{2})/) {	$d[1] --; return timelocal(@d[5,4,3,2,1,0]); }
	# DD.MM.YYYY H:MM:SS
	if(@d = $date =~ m/(\d{2})[\.](\d{2})[\.](\d{4})\s(\d{1,2})[:](\d{2})[:](\d{2})/) { $d[1] --; return timelocal(@d[5,4,3,0,1,2]); }
	# YYYY-MM-DD
	if(@d = $date =~ m/(\d{4})[-](\d{2})[-](\d{2})/) { $d[1] --; return timelocal(0, 0, 0, @d[2,1,0]); }
	return -1;
}

sub dump_cache() {
	print "Saving cache data...";
	my $header = 'myCache';
	my $stati = 0;
	open D, ">", $cache_file;
	binmode D, ":utf8";
	print D $header;
	print D pack "V", length($username);
	print D $username;
	print D pack "V", $last_date;
	foreach my $arti (keys %lastfm_track_playcount) {
		foreach my $titl (keys %{$lastfm_track_playcount{$arti}}) {
			print D pack "V", length($arti);
			print D $arti;
			print D pack "V", length($titl);
			print D $titl;
			print D pack "V", $lastfm_track_playcount{$arti}{$titl};
			print D pack "V", $lastfm_track_playlast{$arti}{$titl};
			$stati += $lastfm_track_playcount{$arti}{$titl};
		}
	}
	close D;
	print "$stati scrobbles\n";
}

sub load_cache() {
	print "Loading cache data...";
	my $header = 'myCache';
	my $ret = ''; my $len = 0;
	my $arti = ''; my $titl = '';
	my $stati = 0;

	open D, "<", $cache_file;
	binmode D, ":utf8";

	read D, $ret, length($header);
	if($ret ne 'myCache') { close D; return undef; }

	read D, $len, 4;
	$len = unpack "V", $len;
	read D, $ret, $len;
	if($ret ne $username) { close D; return undef; }

	read D, $len, 4;
	$cache_last_date = unpack "V", $len;

	while(!eof(D)) {
		read D, $len, 4;
		$len = unpack "V", $len;
		read D, $arti, $len;

		read D, $len, 4;
		$len = unpack "V", $len;
		read D, $titl, $len;

		read D, $len, 4;
		$lastfm_track_playcount{$arti}{$titl} = unpack "V", $len;
		$stati += $lastfm_track_playcount{$arti}{$titl};

		read D, $len, 4;
		$lastfm_track_playlast{$arti}{$titl} = unpack "V", $len;
	}
	close D;
	print "$stati scrobbles\n";
}