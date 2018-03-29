'------------------------------------------------------------
'-              File Name: CodeTest.vb                      -
'------------------------------------------------------------
'-              Written By: Graham Beechum                  -
'-              Written On: 03/28/2018                      -
'------------------------------------------------------------
'-  Program Purpose:                                        -
'-  Calculates the quality of a song from a given album     -
'-  as a function of Zipf's Law where a songs quality is    -
'-  equal to how many times it has been played divided by   -
'-  how many times Zipf's law predicts it will be played.   -
'-  Qi = Fi / Zi                                            -
'------------------------------------------------------------
'-  Global Variable Dictionary (alphabetically):            -
'-  intAmountToShow - Holds the amount of songs to show.    -
'-  dicAlbum - Dictionary holding the name of a song and    -
'-      its play count.                                     -
'-  LINQresults - Object holding results from a linq        -
'-  strFormat - Formatting for string output.               -
'------------------------------------------------------------

Module CodeTest
    Dim dicAlbum As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
    ' Starts as -1, only acceptable answers from user is 1 <= x <= Length of Album.
    Dim intAmountToShow As Integer = -1
    Dim LINQresults As Object
    Dim strFormat As String = "{0,-30}{1,10}{2,10}"


    '------------------------------------------------------------
    '-                  Sub Name: Main                          -
    '------------------------------------------------------------
    '-  Subroutine Purpose:                                     -
    '-  Calls other subroutines to execute the program.         -
    '------------------------------------------------------------
    Sub Main()
        PopulateAlbum()

        GetNumberToShow()

        CalculatePopularity()

        PrintResults()
    End Sub

    '------------------------------------------------------------
    '-                  Sub Name: PopulateAlbum                 -
    '------------------------------------------------------------
    '-  Subroutine Purpose:                                     -
    '-  Populates album with song names and play counts.        -
    '------------------------------------------------------------
    Sub PopulateAlbum()
        ' This can be edited to show whatever albums you wish.
        dicAlbum.Add("One", 30)
        dicAlbum.Add("Two", 30)
        dicAlbum.Add("Three", 15)
        dicAlbum.Add("Four", 25)
    End Sub

    '------------------------------------------------------------
    '-                  Sub Name: GetNumberToShow               -
    '------------------------------------------------------------
    '-  Subroutine Purpose:                                     -
    '-  Prompt the user to enter an amount of songs to be shown.-
    '------------------------------------------------------------
    Sub GetNumberToShow()
        Console.WriteLine("How many songs would you like to see? (1 - " & dicAlbum.Count() & ")")
        ' The user has to enter a number from 1 to the last in the album.
        While Not (intAmountToShow >= 1 And intAmountToShow <= dicAlbum.Count())
            On Error GoTo NotValid
            intAmountToShow = Console.ReadLine()
        End While
        On Error GoTo 0
        Exit Sub

NotValid:
        intAmountToShow = 0
        Resume Next
    End Sub

    '------------------------------------------------------------
    '-                  Sub Name: CalculatePopularity           -
    '------------------------------------------------------------
    '-  Subroutine Purpose:                                     -
    '-  Uses a LINQ to calculate the quality of the song, order -
    '-  them in descending order, and only show how many the    -
    '-  user requested.                                         -
    '------------------------------------------------------------
    Sub CalculatePopularity()
        LINQresults = From Songs In dicAlbum
                      Let Quality = Songs.Value / (1 / GetIndexOfKey(Songs.Key))
                      Order By Quality Descending
                      Select Songs, Quality
                      Take intAmountToShow
    End Sub

    '------------------------------------------------------------
    '-                  Sub Name: PrintResults                  -
    '------------------------------------------------------------
    '-  Subroutine Purpose:                                     -
    '-  Prints the contents of the LINQ result.                 -
    '------------------------------------------------------------
    Sub PrintResults()
        Console.WriteLine(String.Format(strFormat, "Song Name", "Plays", "Quality"))
        Console.WriteLine(String.Format(strFormat, StrDup(9, "-"), StrDup(5, "-"), StrDup(7, "-")))

        For Each song In LINQresults
            Console.WriteLine(String.Format(strFormat,
                                            song.Songs.Key.PadRight(30).Substring(0, 30),
                                            song.Songs.Value,
                                            song.Quality))
        Next
        Console.ReadLine()
    End Sub

    '------------------------------------------------------------
    '-                  Funtion Name: GetIndexOfKey             -
    '------------------------------------------------------------
    '-  Function Purpose:                                       -
    '-  Since I forgot that you can't access a dictionary by    -
    '-  index, this copies the dictionary into an array with    -
    '-  its keys and then finds the index of the key we're      -
    '-  looking for.                                            -
    '------------------------------------------------------------
    '-  Parameter Dictionary (in parameter order):              -
    '-  key - Name of the song we're looking for the index for. -
    '------------------------------------------------------------
    '-  Local Variable Dictionary (alphabetically):             -
    '-  keys() - Array that holds the keys of the dictionary.   -
    '------------------------------------------------------------
    '-  Returns:                                                -
    '-  A one based index of the song we're looking for.        -
    '------------------------------------------------------------
    Function GetIndexOfKey(ByVal key As String) As Integer
        Dim keys(dicAlbum.Count - 1) As String
        dicAlbum.Keys.CopyTo(keys, 0)
        Return (Array.IndexOf(keys, key) + 1)
    End Function
End Module
