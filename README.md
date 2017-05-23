# MidiShow-MidiDownloader-vbs
a simple vbscript file use to download midi(s) from www.midishow.com.

Usage: doubleclick the `.vbs` script and enjoy :D

Here is how it works:

## Function `download` @params: `midid`, `ff`
This function create two objects:

+ `Msxml2.ServerXMLHTTP`, to get binary data from MidiShow's server.
+ `Adodb.Stream`, to write datas in the same directory the `.vbs` file is stored with filename like `<name> - <id>.mid`.

If param `midid` is an url(matches regexp `[A-Za-z]+://www.midishow.com/midi/[^/s]+`), this function will download the midi file; Otherwise, the function `Search` will be called.

##Sub `on_timeout`
Hack. Send an `{Enter}` key to simulation a click on the button `确定/yes`, in order to let the script keep running.

## Function `name` @param: `path(as String)`
This Function returns the name of the midi located at `path`.

## Function `Search` @param: `text(as String)`
This function searches midi(s) containing `text` and show results to the user as well as querying whether the user would like to download that midi or not.

## Function `copy` @param: text(as String), mtl(as Boolean), alt(as Boolean)
This function tries to copy `text` to your clipboard. mtl: Multiline<Boolean> default=`True`, alt: AlertOnFinish<Boolean> default=`False`. Will be used when an error occured in Function `Search`.

## Function `test_connect`
This function check your connections to `www.midishow.com`.
