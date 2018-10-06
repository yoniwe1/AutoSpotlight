using System;
using System.Collections.Generic;
using NAudio.Midi;
using System.IO;
using System.Linq;
using Melanchall.DryWetMidi.Smf;
using Melanchall.DryWetMidi.Smf.Interaction;
using Microsoft.Win32;
using ControlChangeEvent = NAudio.Midi.ControlChangeEvent;
using Excel = Microsoft.Office.Interop.Excel;
using MetaEvent = NAudio.Midi.MetaEvent;
using MidiEvent = NAudio.Midi.MidiEvent;
using MidiFile = Melanchall.DryWetMidi.Smf.MidiFile;
using NoteOnEvent = NAudio.Midi.NoteOnEvent;
using TimeSignatureEvent = NAudio.Midi.TimeSignatureEvent;

namespace AutoSpotlight
{
    internal static class Program
    {
        private static void Main()
        {
            MidiFile midiFile = MidiFile.Read("QuatuorCordes08_Opus59_Num2_Mvt3.mid");
            TempoMap tempoMap = midiFile.GetTempoMap();


            NAudio.Midi.MidiFile myMidi = new NAudio.Midi.MidiFile("QuatuorCordes08_Opus59_Num2_Mvt3.mid");   //create midi file
            myMidi.Events.MidiFileType = 0; //flatten to one track

            try
            {
                GetIntervals(ref myMidi, tempoMap);   //function to retrieve dominant in every interval
                var sW = new StreamWriter("MidiContent.txt");
                sW.Write(myMidi.ToString());
                sW.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
            finally
            {
                Console.WriteLine("Executing finally block.");
            }
        }

        private static string GetPatchName(int patchNumber) 
        {
            return PatchNames[patchNumber];
        }

        private static readonly string[] PatchNames = new string[]  //list of patches
        {
            "Acoustic Grand","Bright Acoustic","Electric Grand","Honky-Tonk","Electric Piano 1","Electric Piano 2","Harpsichord","Clav",
            "Celesta","Glockenspiel","Music Box","Vibraphone","Marimba","Xylophone","Tubular Bells","Dulcimer",
            "Drawbar Organ","Percussive Organ","Rock Organ","Church Organ","Reed Organ","Accoridan","Harmonica","Tango Accordian",
            "Acoustic Guitar(nylon)","Acoustic Guitar(steel)","Electric Guitar(jazz)","Electric Guitar(clean)","Electric Guitar(muted)","Overdriven Guitar","Distortion Guitar","Guitar Harmonics",
            "Acoustic Bass","Electric Bass(finger)","Electric Bass(pick)","Fretless Bass","Slap Bass 1","Slap Bass 2","Synth Bass 1","Synth Bass 2",
            "Violin","Viola","Cello","Contrabass","Tremolo Strings","Pizzicato Strings","Orchestral Strings","Timpani",
            "String Ensemble 1","String Ensemble 2","SynthStrings 1","SynthStrings 2","Choir Aahs","Voice Oohs","Synth Voice","Orchestra Hit",
            "Trumpet","Trombone","Tuba","Muted Trumpet","French Horn","Brass Section","SynthBrass 1","SynthBrass 2",
            "Soprano Sax","Alto Sax","Tenor Sax","Baritone Sax","Oboe","English Horn","Bassoon","Clarinet",
            "Piccolo","Flute","Recorder","Pan Flute","Blown Bottle","Skakuhachi","Whistle","Ocarina",
            "Lead 1 (square)","Lead 2 (sawtooth)","Lead 3 (calliope)","Lead 4 (chiff)","Lead 5 (charang)","Lead 6 (voice)","Lead 7 (fifths)","Lead 8 (bass+lead)",
            "Pad 1 (new age)","Pad 2 (warm)","Pad 3 (polysynth)","Pad 4 (choir)","Pad 5 (bowed)","Pad 6 (metallic)","Pad 7 (halo)","Pad 8 (sweep)",
            "FX 1 (rain)","FX 2 (soundtrack)","FX 3 (crystal)","FX 4 (atmosphere)","FX 5 (brightness)","FX 6 (goblins)","FX 7 (echoes)","FX 8 (sci-fi)",
            "Sitar","Banjo","Shamisen","Koto","Kalimba","Bagpipe","Fiddle","Shanai",
            "Tinkle Bell","Agogo","Steel Drums","Woodblock","Taiko Drum","Melodic Tom","Synth Drum","Reverse Cymbal",
            "Guitar Fret Noise","Breath Noise","Seashore","Bird Tweet","Telephone Ring","Helicopter","Applause","Gunshot"
        };

        private static void GetIntervals(ref NAudio.Midi.MidiFile myMidi, TempoMap tempoMap)   //main function to print the dominant channel on each interval
        {
            //File.WriteAllText(@"output.txt", string.Empty); //empty file
            //var ourStream = File.CreateText(@"output.txt"); //start writing
            var ourStream = new List<string>();
            
            
            //for the Excel output
            
            string output = @"C:\Users\Yoni\Desktop\ThirdStringQuartet.xlsx";
            if (File.Exists(output))
            {
                File.Delete(output);
            }

            Excel.Application oApp;
            Excel.Worksheet oSheet;
            Excel.Workbook oBook;

            oApp = new Excel.Application();
            oBook = oApp.Workbooks.Add();
            oSheet = (Excel.Worksheet)oBook.Worksheets.Item[1];
            // oSheet.Cells[1, 1].EntireRow.Font.Bold = true;
            oSheet.Cells[1, 1].EntireRow.Font.Bold = true;
            oSheet.Cells[1, 1] = "Start Time";

            int lastDominantChannel = 0;
            var currentTempoEvent = new TempoEvent(0,0);
            
            long windowSize = 0;  
            var startEvent = new MidiEvent(myMidi.Events[0][0].AbsoluteTime, myMidi.Events[0][0].Channel,
                myMidi.Events[0][0].CommandCode); //first event in each interval
            IDictionary<int, string> patchToChannel = new Dictionary<int, string>();    //to handle patches 
            IDictionary<int, int>
                mainVolumeToChannel = new Dictionary<int, int>(); //to handle main volume in each channel

            int k = 0;
            List<NoteOnEvent> noteOnCollection2 = new List<NoteOnEvent>();
            while (myMidi.Events[0][k].AbsoluteTime == 0)   //move all noteOns to the end of the 0 absolute time series
            {
                if (myMidi.Events[0][k].CommandCode.ToString() == "NoteOn")
                {
                    var noteOnEvent = (NoteOnEvent)myMidi.Events[0][k];
                    noteOnCollection2.Add(noteOnEvent);
                    myMidi.Events[0].RemoveAt(k);
                    k--;
                }

                k++;
            }

            int s = 0;
            foreach (var noteOnEvent in noteOnCollection2)
            {
                myMidi.Events[0].Insert(k + s, noteOnEvent);
                s++;
            }
            /*for excel!*/
           int counterForExcel = 2;
           var i = 0;
           var endOfFile = false;

            while (true)
            {
                var j = i;
                var exitLoop = false;
                IDictionary<int, int>
                    sumVelocityToChannel = new Dictionary<int, int>(); //to handle sum of velocities per channel
                IDictionary<int, double>
                    meanVelocityToChannel = new Dictionary<int, double>(); //to handle average velocity per channel
                IDictionary<int, int> numOfNoteOnsToChannel = new Dictionary<int, int>();   //to handle number of notes on each channel
                IDictionary<int, double> dominantValue = new Dictionary<int, double>(); //to handle dominant value on each channel
                IDictionary<int, List<NoteOnEvent>>
                    noteOnsToChannel =
                        new Dictionary<int, List<NoteOnEvent>>(); //to handle list of notes on each channel
                List<NoteOnEvent> noteOnCollection = new List<NoteOnEvent>();


                while ((myMidi.Events[0][j].AbsoluteTime - startEvent.AbsoluteTime < windowSize) || windowSize == 0)
                {
                    switch (myMidi.Events[0][j].CommandCode.ToString())
                    {
                        case @"MetaEvent":
                            var metaEvent = (MetaEvent)myMidi.Events[0][j];
                            switch (metaEvent.MetaEventType.ToString())
                            {
                                case @"TimeSignature":
                                    var timeSignatureEvent = (TimeSignatureEvent)metaEvent;
                                    //change window size - change from 8 to 16 or 32 if too fast shifting (from 2 bars window to 4 or 8 bars)
                                    windowSize = (long)Math.Round(myMidi.DeltaTicksPerQuarterNote *
                                                                  ((double)timeSignatureEvent.Numerator /
                                                                   (Math.Pow(2, timeSignatureEvent.Denominator))) * 4); 
                                    exitLoop = true;    //exit interval -> interval size changes
                                    break;

                                case @"EndTrack":
                                    endOfFile = true;   //reached end of midi file
                                    break;

                                default:
                                    break;
                            }

                            break;

                        case @"PatchChange":
                            var patchChangeEvent = (PatchChangeEvent)myMidi.Events[0][j];
                            if (!patchToChannel.ContainsKey(patchChangeEvent.Channel))
                            {
                                patchToChannel.Add(patchChangeEvent.Channel, GetPatchName(patchChangeEvent.Patch));
                            }
                            else
                            {
                                patchToChannel[patchChangeEvent.Channel] = GetPatchName(patchChangeEvent.Patch);
                            }
                            break;

                        case @"ControlChange":
                            var controlChangeEvent = (ControlChangeEvent)myMidi.Events[0][j];
                            if (controlChangeEvent.Controller == MidiController.MainVolume)
                            {
                                if (!mainVolumeToChannel.ContainsKey(controlChangeEvent.Channel))   //if key does not exist - create it
                                {
                                    mainVolumeToChannel.Add(controlChangeEvent.Channel, controlChangeEvent.ControllerValue);
                                }
                                else
                                {
                                    mainVolumeToChannel[controlChangeEvent.Channel] = controlChangeEvent.ControllerValue;
                                    exitLoop = true;
                                }
                            }

                            break;

                        case @"NoteOn":
                            var noteOnEvent = (NoteOnEvent)myMidi.Events[0][j];
                            if (noteOnEvent.Velocity != 0)
                            {
                                if (!noteOnsToChannel.TryGetValue(noteOnEvent.Channel, out noteOnCollection))
                                {
                                    noteOnCollection = new List<NoteOnEvent>();
                                    noteOnsToChannel[noteOnEvent.Channel] = noteOnCollection;
                                    dominantValue[noteOnEvent.Channel] = 0D;
                                }
                                noteOnCollection.Add(noteOnEvent);
                                noteOnsToChannel[noteOnEvent.Channel] = noteOnCollection;
                                numOfNoteOnsToChannel[noteOnEvent.Channel] = noteOnCollection.Count;


                                if (!sumVelocityToChannel.ContainsKey(noteOnEvent.Channel))   //if key does not exist - create it
                                {
                                    sumVelocityToChannel.Add(noteOnEvent.Channel, noteOnEvent.Velocity);
                                }
                                else
                                {
                                    sumVelocityToChannel[noteOnEvent.Channel] += noteOnEvent.Velocity;
                                }
                            }
                            break;

                        default:
                            break;
                    }

                    if (endOfFile)
                    {
                        break;
                    }

                    j++;

                    if (exitLoop)
                    {
                        break;
                    }
                }


                foreach (KeyValuePair<int, int> pair in sumVelocityToChannel)
                {
                    meanVelocityToChannel[pair.Key] =
                        (double) pair.Value /
                        (noteOnsToChannel[pair.Key].Count); //calculate avg velocity on each channel
                }

                var endEvent = new MidiEvent(1,1,0);    //the end event on the current interval

                if (endOfFile)
                {
                    endEvent = new MidiEvent(myMidi.Events[0][j].AbsoluteTime, myMidi.Events[0][j].Channel, myMidi.Events[0][j].CommandCode);
                }
                else
                {
                    endEvent = new MidiEvent(myMidi.Events[0][j-1].AbsoluteTime, myMidi.Events[0][j-1].Channel, myMidi.Events[0][j-1].CommandCode);
                }

                string value = "";  //neccessery for default patch and main volume
                int value1 = 100;

                foreach (KeyValuePair<int, double> pair in meanVelocityToChannel)
                {
                    if (!endOfFile)
                    {
                        if (!dominantValue.ContainsKey(pair.Key))
                        {
                            dominantValue[pair.Key] = 0D;
                        }

                        if (!noteOnsToChannel.ContainsKey(pair.Key))
                        {
                            noteOnsToChannel[pair.Key] = null;
                        }

                        if (!numOfNoteOnsToChannel.ContainsKey(pair.Key))
                        {
                            numOfNoteOnsToChannel[pair.Key] = 0;
                        }

                        if (!mainVolumeToChannel.ContainsKey(pair.Key))
                        {
                            mainVolumeToChannel[pair.Key] = 0;
                        }
                    }
                }

                foreach (KeyValuePair<int, int> pair in mainVolumeToChannel)
                {
                    if (!endOfFile)
                    {
                        if (!dominantValue.ContainsKey(pair.Key))
                        {
                            dominantValue[pair.Key] = 0D;
                        }

                        if (!noteOnsToChannel.ContainsKey(pair.Key))
                        {
                            noteOnsToChannel[pair.Key] = null;
                        }

                        if (!numOfNoteOnsToChannel.ContainsKey(pair.Key))
                        {
                            numOfNoteOnsToChannel[pair.Key] = 0;
                        }

                        if (!meanVelocityToChannel.ContainsKey(pair.Key))
                        {
                            meanVelocityToChannel[pair.Key] = 0D;
                        }
                    }
                }

                foreach (KeyValuePair<int, double> pair in dominantValue)   //default patch and channel volume in case not given in the midi file
                {
                    if (!patchToChannel.TryGetValue(pair.Key, out value))
                    {
                        patchToChannel[pair.Key] = GetPatchName(0);
                    }

                    if (!mainVolumeToChannel.TryGetValue(pair.Key, out value1))
                    {
                        mainVolumeToChannel[pair.Key] = value1;
                    }
                }


               var orderdNumOfNoteOnsToChannel = OrderIDictionary<int>(numOfNoteOnsToChannel);   //sort num of note Ons

                bool isEmpty;
                bool hasNotEmptyValues = false;
                using (var dictionaryEnum = noteOnsToChannel.GetEnumerator())
                {
                    isEmpty = !dictionaryEnum.MoveNext();
                }
                hasNotEmptyValues = noteOnsToChannel
                    .Any(pair => pair.Value != null && pair.Value.Any());

                if (!isEmpty && hasNotEmptyValues) //make sure there is at least one noteOn on what we print
                {
                    IDictionary<int, double> dominant = new Dictionary<int, double>();
                    foreach (KeyValuePair<int, double> pair in dominantValue)   //The formula to change to adapt to different genres
                    {
                        try
                        {
                            if (orderdNumOfNoteOnsToChannel.First().Value -
                                orderdNumOfNoteOnsToChannel.Last().Value != 0)
                            {
                                dominant[pair.Key] = (0.45) * meanVelocityToChannel[pair.Key] +
                                                     (0.2) * (double) mainVolumeToChannel[pair.Key] +
                                                     (0.35) * (double)(((127 * (double)(numOfNoteOnsToChannel[pair.Key] - orderdNumOfNoteOnsToChannel.Last().Value)) / (orderdNumOfNoteOnsToChannel.First().Value - orderdNumOfNoteOnsToChannel.Last().Value))); //calculate dominant value for each channel
                            }
                            else
                            {
                                dominant[pair.Key] = (0.5) * meanVelocityToChannel[pair.Key] +
                                                     (0.5) * (double)mainVolumeToChannel[pair.Key];
                            }

                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                            throw;
                        }
                    }

                    dominantValue = dominant;
                    var orderdDominantValue = OrderIDictionary<double>(dominantValue);   //sort the dominance
               
                    if (orderdDominantValue.First().Key != lastDominantChannel)
                    {
                        ourStream.Add(
                        $"from time {CalcFactor(startEvent.AbsoluteTime, tempoMap)}: Channel {orderdDominantValue.First().Key} ({patchToChannel[orderdDominantValue.First().Key]})");
                            
                        //excel
                        oSheet.Cells[counterForExcel, 1] = CalcFactor(startEvent.AbsoluteTime, tempoMap);
                        oSheet.Cells[counterForExcel, 1].Font.Bold = true;
                        oSheet.Cells[counterForExcel, 2] = orderdDominantValue.First().Key;

                        int v = 2;
                        foreach (KeyValuePair<int, string> pair in patchToChannel)
                        {
                            string str = "";
                            if (!patchToChannel.TryGetValue(pair.Key, out str))
                            {
                                oSheet.Cells[counterForExcel, v] = 0;
                            }
                            else if (pair.Key.Equals(orderdDominantValue.First().Key))
                            {
                                oSheet.Cells[counterForExcel, v] = 1;
                            }
                            else
                            {
                                oSheet.Cells[counterForExcel, v] = 0;
                            }

                            v++;
                        }

                        counterForExcel++;
                        lastDominantChannel = orderdDominantValue.First().Key;
                    }
                }

                startEvent = myMidi.Events[0][j]; //change to j if need big time interval... define start event for next time interval
                 i = j;  //include this if need the big time interval
                if (endOfFile)
                {
                    ourStream.Add($"Ending: {CalcFactor(endEvent.AbsoluteTime, tempoMap)}");
                    oSheet.Cells[counterForExcel, 1].Font.Bold = true;
                    oSheet.Cells[counterForExcel, 1] = CalcFactor(endEvent.AbsoluteTime, tempoMap);
                    oSheet.Cells[counterForExcel, 2] = "Ending";
                    break;
                }
            }

            //ourStream.Close();
            File.WriteAllLines(@"output.txt", ourStream);

            //for excel
            int q = 2;
            foreach (KeyValuePair<int, string> pair in patchToChannel)
            {
                oSheet.Cells[1, q] = $"{pair.Key.ToString()}_({pair.Value})";
                oSheet.Cells[1, q].Font.Bold = true;
                q++;
            }

            oBook.SaveAs(output);
            oBook.Close();
            oApp.Quit();
           
        }

        private static IEnumerable<KeyValuePair<int, T>> OrderIDictionary<T>(IDictionary<int, T> dictionary) //sort dictionary template function
        {
            IOrderedEnumerable<KeyValuePair<int, T>> items = null;
                items = from pair in dictionary
                    orderby pair.Value descending
                    select pair;
            return items;
        }

        private static string CalcFactor(long noteAbsTime, TempoMap tempoMap)  //calculate the factor needed for presenting time in seconds
        {
            MetricTimeSpan metricTime = TimeConverter.ConvertTo<MetricTimeSpan>(noteAbsTime, tempoMap);
            long seconds = 3600 * metricTime.Hours + 60 * metricTime.Minutes + metricTime.Seconds;
            long miliSeconds = metricTime.Milliseconds;
            return $"{seconds}.{miliSeconds}";
        }
    }
}