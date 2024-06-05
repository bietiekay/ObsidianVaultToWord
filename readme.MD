I am using CoPilot at work and I wanted my vault to be indexed with it to be able to ask questions about everything I've noted down in Obsidian.

CoPilot indexes everything on my company OneDrive / Sharepoint and cannot index .MD files today.

Therefore the task/workaround for me is to write a tool which mirrors the whole vault to .docx files to be indexed. It supports update of file contents but not yet deletion. You can run it again to update the contents.

Find the source and binaries for Windows here.

Its a command line tool and takes two parameters. Parameter 1: Source Path, Parameter 2: Target Path.

Maybe it helps other Obsidian Users.