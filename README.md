# iTunesPlaylistExporter
Simple Powershell utility to export iTunes Playlist contents and files

Some of use still have a load of iTunes libraries hanging around and would like to have them on USB or similar media.

If these playlists are exported to .M3U files there is hope.

Use:

1.  Clone the repo to a directory
2.  Open Powershell and navigate to the directory
3.  Run .\Make-Playlist.ps1
4.  The script will start then will launch a dialog for you to browse to the directory where the .m3u files have been saved to

The script runs and parses the playlist.  It creates a new directory, copies all of the files.  It then renames and tags these files and then creates a matching M3U file within.


