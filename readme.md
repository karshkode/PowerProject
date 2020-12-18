PowerProject utilizes a PowerShell file watcher (System.IO.FileSystemWatcher) and monitors a folder of choice to be able to project a PowerPoint presentation by simply dropping a file into it\'92s repository:\

* $watcher.Path: Path to folder where files are watched\
* $LogPath: Path to logs of session\

Stage is utilized for files so that the file library that is watched does not create additional temp files. In this example case, PowerPoint tmp files will not be staged causing indefinite loops and errors within PowerPoint as it will close every time a file is added to the repo\

File types can be specified as to be watched depending on use case of script}