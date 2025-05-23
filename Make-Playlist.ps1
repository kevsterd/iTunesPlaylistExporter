##
## Simple utility to select a M3U playlist folder
## then create folder based on the the name within then
## parse the file and copy the content to it.
##
## Lastly make a replica of the playlist with edited path to the local files
##

## Dependencies
##              - TagLibSharp.dll

## Release Notes
##              - 2025-05-19 - 1.4 - Fix support for paths containing []
##              - 2020-08-22 - 1.3 - Added more code to tag the correct file (eek!) and rename files on copy. 
##              - 2020-08-04 - 1.2 - Added re-tagging code to adjust the MP3 tags in files to match playlist
##              - 2020-08-03 - 1.1 - Update to move if detect folder to an else loop
##                                 - Fix some path issues in PowerShell 5.1
## @kevsterd    - 2020-08-02 - 1.0 - Initial Version

param(
    [bool]$Debug #if true with -Debug switch, will show debug messages
)

#Set the debug preference to continue or not
if ($Debug) {
    $DebugPreference = "Continue"
} else {
    $DebugPreference = "SilentlyContinue"
}

Function Get-Folder($initialDirectory="")
 #Display a Windows selection dialog to allow for easy picking of directories
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder containing playlist files (M3U)"
    $foldername.rootfolder = "MyComputer"
    $foldername.SelectedPath = $initialDirectory

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}

Function Get-MP3MetaData
 #Non .Net method to get file properties (MP3/MP4/more) from the shell
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([Psobject])]
    Param
    (
        [String] [Parameter(Mandatory=$true, ValueFromPipeline=$true)] $Directory
    )

    Begin
    {
        $shell = New-Object -ComObject "Shell.Application"
    }
    Process
    {

        Foreach($Dir in $Directory)
        {
            $ObjDir = $shell.NameSpace($Dir)
            $Files = Get-ChildItem $Dir| Where-Object {$_.Extension -in '.mp3','.mp4'}

            Foreach($File in $Files)
            {
                $ObjFile = $ObjDir.parsename($File.Name)
                $MetaData = @{}
                $MP3 = ($ObjDir.Items()|Where-Object{$_.path -like "*.mp3" -or $_.path -like "*.mp4"})
                $PropertArray = 0,1,2,12,13,14,15,16,17,18,19,20,21,22,27,28,36,220,223
            
                Foreach($item in $PropertArray)
                { 
                    If($ObjDir.GetDetailsOf($ObjFile, $item)) #To avoid empty values
                    {
                        $MetaData[$($ObjDir.GetDetailsOf($MP3,$item))] = $ObjDir.GetDetailsOf($ObjFile, $item)
                    }
                 
                }
            
                New-Object psobject -Property $MetaData |Select-Object *, @{n="Directory";e={$Dir}}, @{n="Fullname";e={Join-Path $Dir $File.Name -Resolve}}, @{n="Extension";e={$File.Extension}}
            }
        }
    }
    End
    {
    }
}

Function Invoke-TagLibSharp
 #Load the TagLibSharp assembly silently
{
    
    [Reflection.Assembly]::LoadFrom( (Resolve-Path ("taglibsharp.dll"))) | Out-Null
}

Function Get-MP3Tags
 #Get the file properties (MP3/MP4/more)
{
    param (
        $SongFile
    )

    #Get Full Path regardless of whatever specified
    $SongFile = (Get-ChildItem $SongFile).FullName

    #$file_name = [System.IO.Path]::GetFileNameWithoutExtension($path_file) 
    #$file_name_array=$file_name.Split("_")

    $MP3Tags = ([TagLib.File]::Create($SongFile).Tag)

    Return $MP3Tags
}

function Set-MP3Tags 
 #Set the file properties (MP3/MP4/more)
{
    param (
        $SongFile,
        $Title,
        $TrackNumber,
        $Artist,
        $Album
    )
   
    #Get Full Path regardless of whatever specified
    $SongFile = (Get-ChildItem $SongFile).FullName

    #Create an object for the song
    $EditSong = ([TagLib.File]::Create($SongFile))

    $EditSong.Tag.Title = $Title
    $EditSong.Tag.Track = $TrackNumber
    $EditSong.Tag.Artists = $Artist
    $EditSong.Tag.AlbumArtists = $Artist
    $EditSong.Tag.Album = $Album
     
    $EditSong.Tag.Disc = ""
    $EditSong.Tag.DiscCount = ""

    try {
        $EditSong.Save()
        }
    catch
        {
            Write-Host "ERROR:  Error saving MP3 tags to $($SongFile)" -ForegroundColor Red
            Write-Host $Error[0]
            Start-Sleep 1
        }
    
}


#Load the assembly
Invoke-TagLibSharp

Write-Host ""
Write-Host "iTunes Playlist Exporter" -ForegroundColor green
Write-Host "------------------------" -ForegroundColor green
Write-Host ""

Write-Host "First you need to export the playlist to a folder from iTunes." -ForegroundColor green
Write-Host "It should be in the format of an M3U.    You can also put multiple playlists in the same folder" -ForegroundColor green
Write-Host ""

Read-Host "-> Press ENTER and then select the folder you exported the playlists to"

#Check if a path was passed in
if ($Path) {
    $folderToParse = $Path
} else {
    #Get the folder to parse
    $folderToParse = Get-Folder
}

Write-Host "Locating Playlist (M3U) files..."  -ForegroundColor Yellow
try {
    $foundPlaylists = Get-ChildItem -Path $folderToParse -File "*.m3u"
    }
catch
    {
        Write-Host "No playlist files found in directory $($folderToParse) " -ForegroundColor Red
        Write-Host "Aborting..."
        Break
    }

#Got here, must of found some.  Now loop each
foreach ($Playlist in $foundPlaylists) {
    Write-Host ""
    Write-Host "- Found $($Playlist)" -ForegroundColor green
    Write-Host ""
   
       
    #First, get the directory and filename deats...
    $PlaylistDir  = $PlayList.DirectoryName
    $PlaylistFile = $PlayList.Name
    
    $NewDir = (($PlaylistFile).Split('.') )[0] 
    
    $FullNewDir = $PlaylistDir + "\" + $NewDir
    $NewPlaylist =  "$($FullNewDir)\$($PlaylistFile)"

    #Define a file counter
    $FileNumber = 0
    
    if (Test-Path $FullNewDir) {
        Write-Host " -- The output directory $($FUllNewDir) already exists.  Skipping ..." -ForegroundColor Red
        }
      
    else {        
        #Create new directory based on the playlist filename
        New-Item -Path $FullNewDir -Type Directory -Force | Out-Null

        #Create a new playlist in the new directory
        New-Item -Path $NewPlaylist -Type File -Force | Out-Null

        #The playlist file to read
        $reader = [io.file]::OpenText($Playlist.FullName)

        #The new improved playlist file
        $writer = [io.file]::CreateText($NewPlaylist)

        while($reader.EndOfStream -ne $true) {
            $line = $reader.Readline()
            if ($line -like '#*') 
                #Copy a standard playlist file line directly.  These always start with hish (#)
                {
                $writer.WriteLine($line);
                }
            else 
                #Muse be a filepath line.  We get the file, copy it to the dir, then set the file path
                {
                    #Increment the counter
                    $FileNumber ++

                    #Create the playlist line.  This is the filename
                    $NewLine = (($Line).Split('\'))[-1]

                    #Remove any track numbers from the file start
                    $NewLine = ($NewLine).TrimStart("0123456789-. ")

                    #Pad the track number with a leading 0 if needed.  Helps sorting and playback.
                    $TrackNumber = "{0:D2}" -f $FileNumber

                    #Preappend the new track number to the start of the file
                    $NewLine = "$($TrackNumber) - $($NewLine)"
                
                    try {
                        #Copy the file to the directory
                        Write-Host " - Copying [$($TrackNumber)] - $($NewLine)" -ForegroundColor Yellow

                        Write-Debug "Source : $($Line)"
                    
                        if (Test-Path -LiteralPath $Line)
                            {
                                Write-Debug "Dest   : $($FullNewDir)\$($NewLine)"
                                Copy-Item -LiteralPath $Line -Destination "$($FullNewDir)\$($NewLine)"
                            }
                        }
                    catch {
                        Write-Host " --- Problem accessing source file ^^^. Skipping..." -ForegroundColor Red
                        Continue
                        }
                
                    #Append a local relative path to the filename
                    $OutLine = ".\" + $NewLine
                    
                    #Write out the track filename to the new playlist file
                    $writer.WriteLine($OutLine);

                    #Get the copied files path
                    try {
                        write-debug "$($FullNewDir)\$($NewLine)"

                        $FindSong = Get-ChildItem -Path "$($FullNewDir)\$($NewLine)" 

                    }
                    catch{
                        Write-Host "ERROR: Error finding file" -ForegroundColor Red
                        Break
                    }
                    
                    #Get the current tags
                    try {
                        $CurrentTags = Get-MP3Tags -SongFile $FindSong
                    }
                    catch {
                        Write-Host "ERROR: Error getting current tags" -ForegroundColor Red
                        
                    }
                    
                    
                    #Set the new tags
                    try {
                        
                        $tmpArtists = $CurrentTags.Artists[0]
                        
                        $tmpTitle = $CurrentTags.Title
                        
                        $tmpFullTitle = "$($tmpArtists) - $($tmpTitle)"
                        
                        Set-MP3Tags -SongFile $FindSong -TrackNumber $TrackNumber -Title $tmpFullTitle -Artist "Various Artists" -Album $NewDir

                        Start-Sleep 1
                    }
                    catch {
                        Write-Host "ERROR: Error setting new tags" -ForegroundColor Red
                        #Break
                    }
                    
                
                }
        }
 
        #Close the files
        $writer.Dispose();
        $reader.Dispose();

        #Tell the world...
        Write-Host ""
        Write-Host "Playlist : $($NewDir) has been created !!" -ForegroundColor green
        Write-Host ""
    }
}

Write-Host ""
Write-Host "Completed" -ForegroundColor Green
