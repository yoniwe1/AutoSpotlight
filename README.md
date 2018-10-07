# AutoSpotlight
AutoSpotlight C# .net framework project, based on NAudio library, to identify the dominant instrument in a given time interval on a midi file piece.

Midi file is a hexadecimal file, built as a sequence of events, each one unique, and NAudio is a library written by Mark Heath which allows us to work fluently with it without having to get down to the resolution of the machine.

The project is divided into two parts: 

[1] The first part is the main algorithm, located on AutoSpotlight/AutoSpotlight algorithm/program.cs. The output of this part is an xls file containing a binary matrix to tell us for what time interval, is a specific instrument the dominant one, as you can see on the path: AutoSpotlight/xls files in every file there.
We start by reading the file twice. Once for the Melanchall library to help us translate the absolute time (in midi ticks per quarter note) to real time. The NAudio library is not free of faults, and one of them is the inability to translate the midi time correctly into real time after multiple tempo events in a file. For this we used the Melanchall library (calcFactor() method). Next, we read the file with the NAudio library and begin the main process. Since all this part is pre-process, on a midi file, time was not an issue for this part. If someone would like to continue to a dynamic live midi input project this should change of course.
next, we flatten the midi file to a serial file, written chronologically by absolute time, instead of clusttered together in track blocks (all events corresponding to a specific midi channel), on line 28. We then form all data structures required, and begin looping on the flattened midi file (line 134). We switch on each possible relavent midi event and act according to our purpose. In our case, we want to calculate the dominance value for each midi channel, using three parameters: 
  * Density of notes in time interval - represented in "numOfNoteOnsToChannel" IDictionary object.
  * Mean Velocity in time interval - represented in "meanVelocityToChannel" IDictionary object.
  * Control channel volume as captured in the midi controller on stage and represented by "mainVolumeToChannel" IDictionary object.
Some other important features which are being used here: TimeSignature Event which helps us calculate the resolution of the interval (1 bar of the piece in our example), EndTrack Event which declares the end of the file in flattened mode, PatchChange Event which symbolizes a change in the instrument playing on a specific channel.
Next, we define default values for each of the parameters, since midi file might not include an initialization for all of them, and the calculation must not fall in such a case. Then, we calculate the dominance level for each of the midi channels, represented in "dominant" IDictionary object, sort them by value and choose the one with higher values to be the channel with 1 value in the binnary table created in the excel file on the output, while all others get a 0.
Note that the weight of each parameter in the dominance formula given on line 358, is chosen intuitively, and that the note density is rescaled to be in the same scale as the other two parameters (0-127).
  


