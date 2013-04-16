Attribute VB_Name = "modArrays"
Option Explicit

' /*********************************************************
' | Name:         mdlArray.bas
' | Description:
' |   -> Package of all array-related procedures I created over the years.
' |   -> Includes many sort algorithms.
' |   ->
' |   -> This code is intellectual property of Philippe Lord.
' |   ->
' |   -> You may use/modify this file as much as you want, as long as this
' |   -> file commented header remains, and more important,
' |   -> that it does not get modified in any possible way.
' |   ->
' |   -> You may find updates of this code at http://Philippe.Lord.MD
' |   -> This code was parsed with Marton's VB Code Formatter v4.
' |   -> That program is a freeware I wrote, available at the above site.
' |   ->
' |   -> If you are hiring personnel, feel free to contact me :)
' |
' | Created:      13 august 2001
' | Author(s) info:
' |   By:         Philippe Lord // Marton
' |   Email:      StromgaldMarton@Hotmail.com
' |   ICQ:        12181387
' | Environment:
' |   -> Created in 1280x1024
' |   -> Arial Narrow 8
' |   -> TAB = 3
' |   -> WinXP 2428
' \*********************************************************
'Notes:
'  -> Binary searchs works only on sorted arrays.
'  -> A hash algorithm can only be applied to a string, explaining the absence of HashSearch on other types than strings.
'  -> HashSearch does not requires anything to be sorted.
'  -> If you add or remove a string from the string array on a hash algorithm, you must ABSOLUTELY rebuild TOTALLY the hash table.
'  -> All indexed search & HashSearch will recreate the index if not supplied (supplied empty).
'  -> Indexed sorts will only sort the index array, thus making the algorithm faster.
'     BUT be warned that it is slower on a long array.
'  -> An hash array is bigger than the original array (about 4 times).
'  -> All sort algorithms includes support for ascending/descending order.
'     However, all functions other than sorting does NOT support descending order.
'  -> Standard QuickSort algorithms are generally fast, but there exists an exception...
'     When the array is *nearly* sorted, QuickSort can be slow (up to 2 times slower).
'     However, the included TriQuickSort algorithm does not suffer from this case, because it combines
'     two sort algorithms, and because it uses 3 medians.
'Efficiency recommendations:
'  (We assume the hash algorithm is based on the full string, not only parts of it.)
'  -> The longer the strings are, the better will a binary search be.
'  -> The bigger the string array, the faster a hash search will be. (comment above has priority over this one)
'  -> If you have under 50 items to sort, use ShellSort.
'  -> If you have over 50 items to sort, use TriQuickSort.
'Functions contained within this .bas file:
'  // Add
'  AddToAnyArray                    ' Adds the data at the nth position.
'  AddToLongArray                   ' Adds the long at the nth position.
'  AddToStringArray                 ' Adds the string at the nth position.
'  AddToSortedAnyArray              ' Adds the data in a sorted array, keeping the array sorted.
'  AddToSortedLongArray             ' Adds the long in a sorted long array, keeping the array sorted.
'  AddToSortedStringArray           ' Adds the string in a sorted string array, keeping the array sorted.
'  AddToIndexedAnyArray             ' Adds the data at the end of the array, keeping the index array sorted.
'  AddToIndexedLongArray            ' Adds the long at the end of the long array, keeping the index array sorted.
'  AddToIndexedStringArray          ' Adds the string at the end of the string array, keeping the index array sorted.
'  // Remove (if one item, array gets erased)
'  RemoveFromAnyArray               ' Removes the nth entry.
'  RemoveFromLongArray              ' Removes the nth long.
'  RemoveFromStringArray            ' Removes the nth string.
'  RemoveFromIndexedAnyArray        ' Removes the nth entry (either array or index), keeping the index array sorted.
'  RemoveFromIndexedLongArray       ' Removes the nth long (either array or index), keeping the index array sorted.
'  RemoveFromIndexedStringArray     ' Removes the nth string (either array or index), keeping the index array sorted.
'  // Hash
'  BuildHashTable                   ' Builds a hash array using sent string array.
'  HashSearch                       ' Returns the position of the searched string on an unsorted string array, using an hash array.
'  // Search (-1 = ERROR_NOT_FOUND)
'  BinarySearchAny                  ' Returns the position of the searched data onto a sorted (ascending) array.
'  BinarySearchLong                 ' Returns the position of the searched long onto a sorted (ascending) long array.
'  BinarySearchString               ' Returns the position of the searched string onto a sorted (ascending) string array.
'  IndexedBinarySearchAny           ' Returns the position of the searched data in an array using a sorted (ascending) index.
'  IndexedBinarySearchLong   (slow) ' Returns the position of the searched long in an array using a sorted (ascending) index.
'  IndexedBinarySearchString        ' Returns the position of the searched string in an array using a sorted (ascending) index.
'  SequentialSearchAnyArray         ' Returns the position of the searched data onto an array.
'  SequentialSearchLongArray        ' Returns the position of the searched long onto a long array.
'  SequentialSearchStringArray      ' Returns the position of the searched string onto a string array.
'  isInAnyArray                     ' Determines if data is in array using a sequential search.
'  isInLongArray                    ' Determines if long is in long array using a sequential search.
'  isInStringArray                  ' Determines if string is in string array using a sequential search.
'  // Sort
'     // < 50 -> ShellSort          ' Efficiency recommandation
'     // >=50 -> TriQuickSort
'  ShellSortAny                     ' Sorts the array.
'  ShellSortLong                    ' Sorts the long array.
'  ShellSortString                  ' Sorts the string array.
'  TriQuickSortAny                  ' Sorts the array.         // TriQuickSort stands for 3-median quicksort algorithm.
'  TriQuickSortLong                 ' Sorts the long array.    // The TriQuickSort algorithm combines with InsertionSort algorithm
'  TriQuickSortString               ' Sorts the string array.  // when the distance gets below 5, which speeds things A LOT (over 40%).
'  IndexedShellSortAny              ' Sorts the index using sent array.
'  IndexedShellSortLong      (slow) ' Sorts the index using sent long array.
'  IndexedShellSortString           ' Sorts the index using sent string array.
'  IndexedTriQuickSortAny           ' Sorts the index using sent array.
'  IndexedTriQuickSortLong   (slow) ' Sorts the index using sent long array.
'  IndexedTriQuickSortString        ' Sorts the index using sent string array.
'  isSortedAnyArray                 ' Determines if the array is sorted.
'  isSortedLongArray                ' Determines if the long array is sorted.
'  isSortedStringArray              ' Determines if the string array is sorted.
'  isSortedIndexedAnyArray          ' Determines if the index is sorted.
'  isSortedIndexedLongArray         ' Determines if the index is sorted.
'  isSortedIndexedStringArray       ' Determines if the index is sorted.
'  // Synchronisation
'  SynchroniseIndexedAnyArray       ' Sorts the array using its index (to get an ascending index).
'  SynchroniseIndexedLongArray      ' Sorts the long array using its index (to get an ascending index).
'  SynchroniseIndexedStringArray    ' Sorts the string array using its index (to get an ascending index).
'  // Copy/Move
'  CopyAnyArray                     ' Copies an array.
'  CopyLongArray                    ' Copies a long array.
'  CopyStringArray                  ' Copies a string array.
'  MoveAnyArray                     ' Moves an array. Source array will be erased (VB function 'Erase').
'  MoveLongArray                    ' Moves a long array. Source array will be erased (VB function 'Erase').
'  MoveStringArray                  ' Moves a string array. Source array will be erased (VB function 'Erase').
'  MergeAnyArray                    ' Merges (combine) 2 arrays. Source array will be erased (VB function 'Erase').
'  MergeLongArray                   ' Merges (combine) 2 long arrays. Source array will be erased (VB function 'Erase').
'  MergeStringArray                 ' Merges (combine) 2 string arrays. Source array will be erased (VB function 'Erase').
'  // Save/Load
'  SaveLongArray                    ' Dumps a long array in a string.
'  SaveStringArray                  ' Dumps a string array in a string.
'  LoadLongArray                    ' Rebuilds a long array from a string dump.
'  LoadStringArray                  ' Rebuilds a string array from a string dump.
'  // Others
'  CreateArray                      ' Returns an array of the type of the first sent argument.
'  DebugDumpArray                   ' MsgBox an array. Use for debugging.
'  ReverseAnyArray                  ' Reverses (inverts) an array.
'  ReverseLongArray                 ' Reverses (inverts) a long array.
'  ReverseStringArray               ' Reverses (inverts) a string array.
'Editorial on the TriQuickSort algorithm - Why is TriQuickSort so fast ?
'  Since the TriQuickSort algorithm is in no way a standard sort algorithm, I will try and explain it here.
'  First, I must say that the main idea started from Sun Microsystems, in java source code form. I found
'  Sun's source code after a search on the internet for the 'fastest' sort algorithm (considering a uniprocessor
'  configuration and a nearly-sorted OR totally unsorted array). I compared the ones that performed the best,
'  and soon enough stumbled upon this one (Sun's one). Generally speaking, it was a 3-median QuickSort, a little
'  twinked, of course. The 3-median QuickSort has the advantage of not suffering standard 2-median QuickSort's
'  problems conserning nearly-sorted arrays  (side-note: ever tried sorting a nearly-sorted array using QuickSort?
'  In case you didn't, let me tell you it gets REALLY slow, it can get MUCH slower than bubblesort in certain cases !).
'  It performed very well, but there was a side-note suggesting using a second algorithm when the number of
'  iterations came low (under 10). I though about it, and understood why they suggested that. First, for those
'  who doesn't know how (generally speaking) a QuickSort works, I'll resume it shortly.
'
'     QuickSort is a recursive algorithm (thus eating lots of RAM) which splits in 2 the array,
'     moving the highest ones the right side, and the lower ones the left side, but without sorting either data
'     on the left or right side, all it does is putting all the lowest on the left and the highest on the right.
'     Then, to sort, it calls itself back (recursively) on the left side, and the right side.
'     It continues like this until everything gets sorted. Now there's 2 major problems with this.
'     One is memory usage, and the second is inefficacity (slow) when the borders are close
'     (when 'low' in the recursive tree) (just keep in mind I'm not going into details).
'
'  So now you should understand why I did another version of Sun's sort algorithm ;) I started up the algorithm
'  by porting java source to VB, which led to some difficulties due to the fact that VB does not 'short-cuts'
'  expressions evaluations, making it crash thru a pure porting.
'
'     ex: While (i - 1 >= LBound(sArray)) And (sArray(i - 1) > sTemp)      ' sArray(i - 1) CRASH !!!
'
'  Ok, this was easily fixed, but should give you a small idea of what had to be done. After porting their
'  3-median QuickSort, I made it stop when the delta (difference) of the 2 bounds came under 10, like
'  suggested by Sun Microsystems. Now, if you understood my explanation of the QuickSort algorithm, you
'  should understand too that stopping the process at delta 10 means all you have to do after QuickSorting
'  is to sort each sections of 10, without needing to do any compares with anything else other than the 10
'  entries you're processing. Imagine just that the cutted-QuickSort sorts generally, but you need to finish
'  the work off by processing packets of 10 entries.
'
'  But I must clarify one point.
'
'  Stopping the recursion tree using a delta 10 does not means IN ANY WAY that you're goin to have sections
'  exactly of 10 'well-placed' entries. In fact, if you think well about the problem, and if you understand
'  well the QuickSort algorithm, it means that your sections can vary from your input delta (10)
'  up to 2x delta -1 (19). If you don't understand the previous remark, either trace the QuickSort's code, or read back.
'
'  So what does that means? Well, it means my previous statement (3 paragraphs above) we're not true if you consider
'  10 to be the only valid delta. Consider either a range from 10 to 19. Now read back 3 paragraphs above ;)
'
'  So that was my first idea...sorting each sections individually.
'
'  I searched for the fastest algorithm for processing small arrays, and had in mind to call it n times, where
'  n equals the number of sections. You must keep in mind that to have a good sort algorithm working
'  on 10-19 entries it means your algorithm have to be as simple as possible, because you cannot even afford
'  to do simple mathematical operations. You just need something simple. And fast ;)
'
'  I though of bubblesort first, but later I came up with a similar algorithm, which has the
'  advantage of not being tied to work with a fixed number of entry (because for god's sake i would never let
'  bubblesort the whole array down !). But, since it's roots are based on bubblesort's algorithm,
'  for it to be effective you must keep the delta very low, under 10. That algorithm I'm talking about is
'  called InsertionSort, which sadly was not designed by me. I found InsertionSort to be the perfect algorithm
'  to continue the cutted-QuickSort's job. I'll copy-paste InsertionSort's algorithm below, it's pretty simple.
'  But, like I said earlier, delta 10 (which gives us a 10-19 section's range) would be like saying: Hey, let's
'  give out the main job to InsertionSort (which is normally slow, but in our case it gets VERY fast), which,
'  BTW, is VERY stupid. You can guess I lowered down the QuickSort's delta. If you look at TriQuickSort's source
'  code, you'll notice the parameter iSplit is the delta I'm talking about. I've put a default value of 4 for it,
'  which gives a sections ranging from 4 to 7 in length, which gives very good results. I do not recommend you
'  put a lower value to it, because QuickSort would eat up too much memory AND starts getting slow. If you put
'  higher than 4, the reverse happens...you get a MUCH lower performance because InsertionSort starts bottlenecking
'  a little too much.
'
'      Private Sub InsertionSortAny(ByRef vArray As Variant, ByVal iMin As Long, ByVal iMax As Long)
'         Dim i     As Long
'         Dim j     As Long
'         Dim vTemp As Variant
'
'         For i = iMin + 1 To iMax
'            vTemp = vArray(i)
'            j = i
'
'            Do While j > iMin
'               If vArray(j - 1) <= vTemp Then Exit Do
'
'               vArray(j) = vArray(j - 1)
'               j = j - 1
'            Loop
'
'            vArray(j) = vTemp
'         Next i
'      End Sub
'
'SYSTEM:
'-------
' -> P3 650e overclocked to 845MHz
' -> 384M RAM PC 133
' -> WinXP 2428
'
'BENCHMARKS:
'-----------
'
'(All benchmarks are made on an array of 10 000 strings having a length of 100 characters ranging from A to Z)
'
'(All results in seconds)
'
'
'Using Non-CopyMemory optimized sort algorythm
'------------------------------------------------AVG-------%-----
'BubbleSort   125.8012  124.6600  125.4101       125.2904  -59421
'ShellSort    0.5310    0.5325    0.5106         0.5247    -149.3
'QuickSort    0.2404    0.2481    0.2425         0.2437    -15.77
'TriQuickSort 0.2107    0.2089    0.2120         0.2105    0.0000
'
'Using CopyMemory optimized sort algorythm
'------------------------------------------------AVG-------%-----
'BubbleSort   59.9765   59.3455   59.3642        59.5621   -43471
'ShellSort    0.3017    0.3121    0.2999         0.3046    -122.8
'QuickSort    0.1812    0.1788    0.1806         0.1802    -31.82
'TriQuickSort 0.1309    0.1383    0.1408         0.1367    0.0000
'
'Using CopyMemory optimized sort algorythm on already sorted string array
'------------------------------------------------AVG-------%-------------
'BubbleSort   24.1941   24.1231   24.1744        24.1639   -32731
'ShellSort    0.1215    0.1100    0.1188         0.1167    -58.56
'QuickSort    0.0892    0.1011    0.1000         0.0968    -31.15
'TriQuickSort 0.0796    0.0709    0.0702         0.0736    0.0000
'
'Using CopyMemory optimized sort algorythm on nearly-sorted string array
'------------------------------------------------AVG-------%------------
'After sorting, we do this (below), then we benchmark the following sort.
'   For i = 0 To n - 1 Step 3
'      SwapStrings sArray(i), sArray(i + 1)
'   Next i
'
'BubbleSort   24.1350   24.1254   24.1764        24.1456   -27911
'ShellSort    0.1328    0.1218    0.1187         0.1244    -44.32     ' notice that ShellSort beats QuickSort here in some cases.
'QuickSort    0.1228    0.1194    0.1181         0.1201    -39.33
'TriQuickSort 0.0796    0.0795    0.0994         0.0862    0.0000
'
'
'RESULTS:
'--------
'
'ALGORYTHM------% SLOWER--
'-------------------------
'BubbleSort     -40884
'ShellSort      -93.75
'QuickSort      -29.52
'TriQuickSort   0.0000
' CopyMemory, my best friend ;)
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (ByRef lpDest As Any, _
                                       ByRef lpSource As Any, _
                                       ByVal iLen As Long)

Private Const ERROR_NOT_FOUND As Long = &H80000000 ' DO NOT CHANGE, for internal usage only !

Public Enum SortOrder
   SortAscending = 0
   SortDescending = 1
End Enum

Public Enum RemoveFrom
   RemoveArray = 0
   RemoveIndex = 1
End Enum

#Const mdlArray_Loaded = True ' DO NOT EDIT !!!
#Const mdlMarton_Loadable = True
' /////////
' // Add //
' /////////
Public Sub AddToAnyArray(ByRef vArray As Variant, _
                         ByVal vToAdd As Variant, _
                         Optional ByVal iPos As Long = -1)
   Dim i       As Long
   Dim iUBound As Long
   Dim j       As Long

   If Not IsArray(vArray) Then Exit Sub
   iUBound = UBound(vArray)

   If iUBound = -1 Then vArray = Array(vToAdd): Exit Sub

   ' if invalid iPos
   If (iPos > iUBound) Or (iPos = -1) Then iPos = iUBound + 1    ' +1 because we can add array past it's end
   If iPos < 0 Then iPos = 0
   iUBound = iUBound + 1
   ReDim Preserve vArray(iUBound)
   j = iPos + 1

   For i = iUBound To j Step -1
      vArray(i) = vArray(i - 1)
   Next

   vArray(iPos) = vToAdd
End Sub

Public Sub AddToLongArray(ByRef iArray() As Long, _
                          ByVal iToAdd As Long, _
                          Optional ByVal iPos As Long = -1)
   Dim iUBound As Long
   iUBound = UBound(iArray)

   If iUBound = -1 Then
      ReDim iArray(0)
      iArray(0) = iToAdd
      Exit Sub

   End If

   ' if adding at the end
   If (iPos > iUBound) Or (iPos = -1) Then
      ReDim Preserve iArray(iUBound + 1)
      iArray(iUBound + 1) = iToAdd
      Exit Sub

   End If

   If iPos < 0 Then iPos = 0
   iUBound = iUBound + 1
   ReDim Preserve iArray(iUBound)
   CopyMemory iArray(iPos + 1), iArray(iPos), (iUBound - LBound(iArray) - iPos) * Len(iArray(iPos))
   iArray(iPos) = iToAdd
End Sub

Public Sub AddToStringArray(ByRef sArray() As String, _
                            ByVal sStringToAdd As String, _
                            Optional ByVal iPos As Long = -1)
   Dim iUBound As Long
   Dim iTemp   As Long
   iUBound = UBound(sArray)

   If iUBound = -1 Then
      ReDim sArray(0)
      sArray(0) = sStringToAdd
      Exit Sub

   End If

   ' if adding at the end
   If (iPos > iUBound) Or (iPos = -1) Then
      ReDim Preserve sArray(iUBound + 1)
      sArray(iUBound + 1) = sStringToAdd
      Exit Sub

   End If

   If iPos < 0 Then iPos = 0
   iUBound = iUBound + 1
   ReDim Preserve sArray(iUBound)
   CopyMemory ByVal VarPtr(sArray(iPos + 1)), ByVal VarPtr(sArray(iPos)), (iUBound - iPos) * 4
   iTemp = 0 ' view this as String(4, Chr(0)) or a NULL value
   CopyMemory ByVal VarPtr(sArray(iPos)), iTemp, 4
   sArray(iPos) = sStringToAdd
End Sub

Public Sub AddToSortedAnyArray(ByRef vArray As Variant, _
                               ByVal vToAdd As Variant)
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iMiddle As Long
   Dim i       As Long

   If Not IsArray(vArray) Then Exit Sub
   iLBound = LBound(vArray)
   iUBound = UBound(vArray)

   ' first, we check the bounds
   If vToAdd <= vArray(iLBound) Then AddToAnyArray vArray, vToAdd, iLBound: Exit Sub
   If vToAdd >= vArray(iUBound) Then AddToAnyArray vArray, vToAdd, iUBound + 1: Exit Sub

   Do
      iMiddle = (iLBound + iUBound) \ 2

      If vArray(iMiddle) = vToAdd Then
         Exit Do
      ElseIf vArray(iMiddle) < vToAdd Then
         iLBound = iMiddle + 1
      Else
         iUBound = iMiddle - 1
      End If

   Loop Until iLBound > iUBound

   iLBound = LBound(vArray)
   iUBound = UBound(vArray)

   For i = iMiddle To iLBound Step -1

      If vArray(i) <= vToAdd Then Exit For
   Next

   If vArray(i) = vToAdd Then AddToAnyArray vArray, vToAdd, i: Exit Sub

   For i = i + 1 To iUBound

      If vArray(i) >= vToAdd Then AddToAnyArray vArray, vToAdd, i: Exit Sub
   Next

End Sub

Public Sub AddToSortedLongArray(ByRef iArray() As Long, _
                                ByVal iToAdd As Long)
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iMiddle As Long
   Dim i       As Long
   iLBound = LBound(iArray)
   iUBound = UBound(iArray)

   ' first, we check the bounds
   If iToAdd <= iArray(iLBound) Then AddToLongArray iArray, iToAdd, iLBound: Exit Sub
   If iToAdd >= iArray(iUBound) Then AddToLongArray iArray, iToAdd, iUBound + 1: Exit Sub

   Do
      iMiddle = (iLBound + iUBound) \ 2

      If iArray(iMiddle) = iToAdd Then
         Exit Do
      ElseIf iArray(iMiddle) < iToAdd Then
         iLBound = iMiddle + 1
      Else
         iUBound = iMiddle - 1
      End If

   Loop Until iLBound > iUBound

   iLBound = LBound(iArray)
   iUBound = UBound(iArray)

   For i = iMiddle To iLBound Step -1

      If iArray(i) <= iToAdd Then Exit For
   Next

   If iArray(i) = iToAdd Then AddToLongArray iArray, iToAdd, i: Exit Sub

   For i = i + 1 To iUBound

      If iArray(i) >= iToAdd Then AddToLongArray iArray, iToAdd, i: Exit Sub
   Next

End Sub

Public Sub AddToSortedStringArray(ByRef sArray() As String, _
                                  ByVal sToAdd As String)
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iMiddle As Long
   Dim i       As Long
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)

   ' first, we check the bounds
   If sToAdd <= sArray(iLBound) Then AddToStringArray sArray, sToAdd, iLBound: Exit Sub
   If sToAdd >= sArray(iUBound) Then AddToStringArray sArray, sToAdd, iUBound + 1: Exit Sub

   Do
      iMiddle = (iLBound + iUBound) \ 2

      If sArray(iMiddle) = sToAdd Then
         Exit Do
      ElseIf sArray(iMiddle) < sToAdd Then
         iLBound = iMiddle + 1
      Else
         iUBound = iMiddle - 1
      End If

   Loop Until iLBound > iUBound

   iLBound = LBound(sArray)
   iUBound = UBound(sArray)

   For i = iMiddle To iLBound Step -1

      If sArray(i) <= sToAdd Then Exit For
   Next

   If sArray(i) = sToAdd Then AddToStringArray sArray, sToAdd, i: Exit Sub

   For i = i + 1 To iUBound

      If sArray(i) >= sToAdd Then AddToStringArray sArray, sToAdd, i: Exit Sub
   Next

End Sub

Public Sub AddToIndexedAnyArray(ByRef vArray As Variant, _
                                ByRef iIndexArray() As Long, _
                                ByVal vToAdd As Variant)
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iMiddle As Long
   Dim i       As Long

   If Not IsArray(vArray) Then Exit Sub
   AddToAnyArray vArray, vToAdd  ' this adds at the end
   iLBound = LBound(vArray)
   iUBound = UBound(vArray)

   ' first, we check the bounds
   If vToAdd <= vArray(iIndexArray(iLBound)) Then AddToLongArray iIndexArray, iUBound, iLBound: Exit Sub
   If vToAdd >= vArray(iIndexArray(iUBound - 1)) Then AddToLongArray iIndexArray, iUBound: Exit Sub

   Do
      iMiddle = (iLBound + iUBound) \ 2

      If vArray(iIndexArray(iMiddle)) = vToAdd Then
         Exit Do
      ElseIf vArray(iIndexArray(iMiddle)) < vToAdd Then
         iLBound = iMiddle + 1
      Else
         iUBound = iMiddle - 1
      End If

   Loop Until iLBound > iUBound

   iLBound = LBound(vArray)
   iUBound = UBound(vArray)

   For i = iMiddle To iLBound Step -1

      If vArray(iIndexArray(i)) <= vToAdd Then Exit For
   Next

   For i = i To iUBound

      If vArray(iIndexArray(i)) >= vToAdd Then AddToLongArray iIndexArray, iUBound, i: Exit Sub
   Next

End Sub

Public Sub AddToIndexedLongArray(ByRef iArray() As Long, _
                                 ByRef iIndexArray() As Long, _
                                 ByVal iToAdd As Long)
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iMiddle As Long
   Dim i       As Long
   AddToLongArray iArray, iToAdd  ' this adds at the end
   iLBound = LBound(iArray)
   iUBound = UBound(iArray)

   ' first, we check the bounds
   If iToAdd <= iArray(iIndexArray(iLBound)) Then AddToLongArray iIndexArray, iUBound, iLBound: Exit Sub
   If iToAdd >= iArray(iIndexArray(iUBound - 1)) Then AddToLongArray iIndexArray, iUBound: Exit Sub

   Do
      iMiddle = (iLBound + iUBound) \ 2

      If iArray(iIndexArray(iMiddle)) = iToAdd Then
         Exit Do
      ElseIf iArray(iIndexArray(iMiddle)) < iToAdd Then
         iLBound = iMiddle + 1
      Else
         iUBound = iMiddle - 1
      End If

   Loop Until iLBound > iUBound

   iLBound = LBound(iArray)
   iUBound = UBound(iArray)

   For i = iMiddle To iLBound Step -1

      If iArray(iIndexArray(i)) <= iToAdd Then Exit For
   Next

   For i = i To iUBound

      If iArray(iIndexArray(i)) >= iToAdd Then AddToLongArray iIndexArray, iUBound, i: Exit Sub
   Next

End Sub

Public Sub AddToIndexedStringArray(ByRef sArray() As String, _
                                   ByRef iIndexArray() As Long, _
                                   ByVal sToAdd As String)
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iMiddle As Long
   Dim i       As Long
   AddToStringArray sArray, sToAdd  ' this adds at the end
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)

   ' first, we check the bounds
   If sToAdd <= sArray(iIndexArray(iLBound)) Then AddToLongArray iIndexArray, iUBound, iLBound: Exit Sub
   If sToAdd >= sArray(iIndexArray(iUBound - 1)) Then AddToLongArray iIndexArray, iUBound: Exit Sub

   Do
      iMiddle = (iLBound + iUBound) \ 2

      If sArray(iIndexArray(iMiddle)) = sToAdd Then
         Exit Do
      ElseIf sArray(iIndexArray(iMiddle)) < sToAdd Then
         iLBound = iMiddle + 1
      Else
         iUBound = iMiddle - 1
      End If

   Loop Until iLBound > iUBound

   iLBound = LBound(sArray)
   iUBound = UBound(sArray)

   For i = iMiddle To iLBound Step -1

      If sArray(iIndexArray(i)) <= sToAdd Then Exit For
   Next i

   For i = i To iUBound

      If sArray(iIndexArray(i)) >= sToAdd Then AddToLongArray iIndexArray, iUBound, i: Exit Sub
   Next i

End Sub

' ////////////
' // Remove //
' ////////////
Public Sub RemoveFromAnyArray(ByRef vArray As Variant, _
                              Optional ByVal iPos As Long = -1)
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long

   If Not IsArray(vArray) Then Exit Sub
   iLBound = LBound(vArray)
   iUBound = UBound(vArray)

   ' if we only have one element in array
   If (iUBound = -1) Or (iUBound - iLBound = 0) Then Erase vArray: Exit Sub

   ' if invalid iPos
   If (iPos > iUBound) Or (iPos = -1) Then iPos = iUBound
   If iPos < iLBound Then iPos = iLBound
   If iPos = iUBound Then ReDim Preserve vArray(iUBound - 1): Exit Sub

   For i = iPos + 1 To iUBound
      vArray(i - 1) = vArray(i)
   Next i

   ReDim Preserve vArray(iUBound - 1)
End Sub

Public Sub RemoveFromLongArray(ByRef iArray() As Long, _
                               Optional ByVal iPos As Long = -1)
   Dim iLBound As Long
   Dim iUBound As Long
   iLBound = LBound(iArray)
   iUBound = UBound(iArray)

   ' if we only have one element in array
   If (iUBound = -1) Or (iUBound - iLBound = 0) Then Erase iArray: Exit Sub

   ' if invalid iPos
   If (iPos > iUBound) Or (iPos = -1) Then iPos = iUBound
   If iPos < iLBound Then iPos = iLBound
   If iPos = iUBound Then ReDim Preserve iArray(iUBound - 1): Exit Sub
   CopyMemory iArray(iPos), iArray(iPos + 1), (iUBound - iLBound - iPos) * Len(iArray(iPos))
   ReDim Preserve iArray(iUBound - 1)
End Sub

Public Sub RemoveFromStringArray(ByRef sArray() As String, _
                                 Optional ByVal iPos As Long = -1)
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iTemp   As Long
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)

   ' if we only have one element in array
   If (iUBound = -1) Or (iUBound - iLBound = 0) Then Erase sArray: Exit Sub

   ' if invalid iPos
   If (iPos > iUBound) Or (iPos = -1) Then iPos = iUBound
   If iPos < iLBound Then iPos = iLBound
   If iPos = iUBound Then ReDim Preserve sArray(iUBound - 1): Exit Sub
   iTemp = StrPtr(sArray(iPos))
   CopyMemory ByVal VarPtr(sArray(iPos)), ByVal VarPtr(sArray(iPos + 1)), (iUBound - iPos) * 4
   ' we do this to have VB unalloc the string to evade memory leaks
   CopyMemory ByVal VarPtr(sArray(iUBound)), iTemp, 4
   ReDim Preserve sArray(iUBound - 1)
End Sub

Public Sub RemoveFromIndexedAnyArray(ByRef vArray As Variant, _
                                     ByRef iIndexArray() As Long, _
                                     Optional ByVal iPos As Long = -1, _
                                     Optional ByVal RemoveFrom As RemoveFrom = RemoveIndex)
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iTemp   As Long
   Dim iPos2   As Long

   If Not IsArray(vArray) Then Exit Sub
   iLBound = LBound(vArray)
   iUBound = UBound(vArray)

   ' if we only have one element in array
   If (iUBound = -1) Or (iUBound - iLBound = 0) Then Erase vArray: Erase iIndexArray: Exit Sub

   ' if invalid iPos
   If (iPos > iUBound) Or (iPos = -1) Then iPos = iUBound
   If iPos < iLBound Then iPos = iLBound
   iTemp = IIf(RemoveFrom = RemoveArray, iPos, iIndexArray(iPos))
   iPos2 = 0

   For i = iLBound To iUBound

      If iIndexArray(i) > iTemp Then
         iIndexArray(i) = iIndexArray(i) - 1
      ElseIf iIndexArray(i) = iTemp Then
         iPos2 = i
      End If

   Next i

   RemoveFromAnyArray vArray, iTemp
   RemoveFromLongArray iIndexArray, IIf(RemoveFrom = RemoveArray, iPos2, iPos)
End Sub

Public Sub RemoveFromIndexedLongArray(ByRef iArray() As Long, _
                                      ByRef iIndexArray() As Long, _
                                      Optional ByVal iPos As Long = -1, _
                                      Optional ByVal RemoveFrom As RemoveFrom = RemoveIndex)
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iTemp   As Long
   Dim iPos2   As Long
   iLBound = LBound(iArray)
   iUBound = UBound(iArray)

   ' if we only have one element in array
   If (iUBound = -1) Or (iUBound - iLBound = 0) Then Erase iArray: Erase iIndexArray: Exit Sub

   ' if invalid iPos
   If (iPos > iUBound) Or (iPos = -1) Then iPos = iUBound
   If iPos < iLBound Then iPos = iLBound
   iTemp = IIf(RemoveFrom = RemoveArray, iPos, iIndexArray(iPos))
   iPos2 = 0

   For i = iLBound To iUBound

      If iIndexArray(i) > iTemp Then
         iIndexArray(i) = iIndexArray(i) - 1
      ElseIf iIndexArray(i) = iTemp Then
         iPos2 = i
      End If

   Next i

   RemoveFromLongArray iArray, iTemp
   RemoveFromLongArray iIndexArray, IIf(RemoveFrom = RemoveArray, iPos2, iPos)
End Sub

Public Sub RemoveFromIndexedStringArray(ByRef sArray() As String, _
                                        ByRef iIndexArray() As Long, _
                                        Optional ByVal iPos As Long = -1, _
                                        Optional ByVal RemoveFrom As RemoveFrom = RemoveIndex)
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iTemp   As Long
   Dim iPos2   As Long
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)

   ' if we only have one element in array
   If (iUBound = -1) Or (iUBound - iLBound = 0) Then Erase sArray: Erase iIndexArray: Exit Sub

   ' if invalid iPos
   If (iPos > iUBound) Or (iPos = -1) Then iPos = iUBound
   If iPos < iLBound Then iPos = iLBound
   iTemp = IIf(RemoveFrom = RemoveArray, iPos, iIndexArray(iPos))
   iPos2 = 0

   For i = iLBound To iUBound

      If iIndexArray(i) > iTemp Then
         iIndexArray(i) = iIndexArray(i) - 1
      ElseIf iIndexArray(i) = iTemp Then
         iPos2 = i
      End If

   Next i

   RemoveFromStringArray sArray, iTemp
   RemoveFromLongArray iIndexArray, IIf(RemoveFrom = RemoveArray, iPos2, iPos)
End Sub

' //////////
' // Hash //
' //////////
Public Sub BuildHashTable(ByRef sArray() As String, _
                          ByRef iHashArray() As Long)
   Dim i        As Long ' Loop Counter
   Dim iLBound  As Long
   Dim iUBound  As Long
   Dim iUBound2 As Long
   Dim iMax     As Long
   Dim iIndex   As Long
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)
   iMax = (iUBound + 1) * 4
   ReDim iHashArray(0 To iMax - 1) As Long
   iUBound2 = UBound(iHashArray)

   For i = LBound(iHashArray) To iUBound2
      iHashArray(i) = ERROR_NOT_FOUND
   Next

   For i = iLBound To iUBound
      iIndex = GetFastXorHash(sArray(i)) Mod iMax

      Do Until iHashArray(iIndex) = ERROR_NOT_FOUND ' remember the hash array is 4 time bigger than the string array, thus this CANNOT be an infinite loop
         iIndex = (iIndex + 1) Mod iMax
      Loop

      iHashArray(iIndex) = i
   Next i

End Sub

Public Function HashSearch(ByRef sArray() As String, _
                           ByRef iHashArray() As Long, _
                           ByVal sFind As String) As Long
   Dim i           As Long
   Dim iMax        As Long
   Dim bInitialize As Boolean
   ' create the hash array if necessary
   bInitialize = False

   If UBound(iHashArray) = -1 Then bInitialize = True Else If iHashArray(LBound(iHashArray)) = iHashArray(UBound(iHashArray)) Then bInitialize = True
   If bInitialize = True Then BuildHashTable sArray, iHashArray
   iMax = UBound(iHashArray) + 1
   i = GetFastXorHash(sFind) Mod iMax

   Do Until iHashArray(i) = ERROR_NOT_FOUND

      If sArray(iHashArray(i)) = sFind Then HashSearch = iHashArray(i): Exit Function
      i = (i + 1) Mod iMax
   Loop

   HashSearch = -1
End Function

' ////////////
' // Search //
' ////////////
Public Function BinarySearchAny(ByRef vArray As Variant, _
                                ByVal vFind As Variant) As Long
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iMiddle As Long

   If Not IsArray(vArray) Then Exit Function
   iLBound = LBound(vArray)
   iUBound = UBound(vArray)

   Do
      iMiddle = (iLBound + iUBound) \ 2

      If vArray(iMiddle) = vFind Then
         BinarySearchAny = iMiddle
         Exit Function

      ElseIf vArray(iMiddle) < vFind Then
         iLBound = iMiddle + 1
      Else
         iUBound = iMiddle - 1
      End If

   Loop Until iLBound > iUBound

   BinarySearchAny = -1
End Function

Public Function BinarySearchLong(ByRef iArray() As Long, _
                                 ByVal iFind As Long) As Long
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iMiddle As Long
   iLBound = LBound(iArray)
   iUBound = UBound(iArray)

   Do
      iMiddle = (iLBound + iUBound) \ 2

      If iArray(iMiddle) = iFind Then
         BinarySearchLong = iMiddle
         Exit Function

      ElseIf iArray(iMiddle) < iFind Then
         iLBound = iMiddle + 1
      Else
         iUBound = iMiddle - 1
      End If

   Loop Until iLBound > iUBound

   BinarySearchLong = -1
End Function

Public Function BinarySearchString(ByRef sArray() As String, _
                                   ByVal sFind As String) As Long
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iMiddle As Long
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)

   Do
      iMiddle = (iLBound + iUBound) \ 2

      If sArray(iMiddle) = sFind Then
         BinarySearchString = iMiddle
         Exit Function

      ElseIf sArray(iMiddle) < sFind Then
         iLBound = iMiddle + 1
      Else
         iUBound = iMiddle - 1
      End If

   Loop Until iLBound > iUBound

   BinarySearchString = -1
End Function

Public Function IndexedBinarySearchAny(ByRef vArray As Variant, _
                                       ByRef iIndexArray() As Long, _
                                       ByVal vFind As Variant) As Long
   Dim iLBound     As Long
   Dim iUBound     As Long
   Dim iMiddle     As Long
   Dim bInitialize As Boolean

   If Not IsArray(vArray) Then Exit Function
   iLBound = LBound(vArray)
   iUBound = UBound(vArray)
   'initialize the index array if necessary
   bInitialize = False

   If UBound(iIndexArray) = -1 Then bInitialize = True Else If iIndexArray(LBound(iIndexArray)) = 0 And iIndexArray(UBound(iIndexArray)) = 0 Then bInitialize = True
   If bInitialize = True Then CreateIndex iIndexArray, vArray

   Do
      iMiddle = (iLBound + iUBound) \ 2

      If vArray(iIndexArray(iMiddle)) = vFind Then
         IndexedBinarySearchAny = iIndexArray(iMiddle)
         Exit Function

      ElseIf vArray(iIndexArray(iMiddle)) < vFind Then
         iLBound = iMiddle + 1
      Else
         iUBound = iMiddle - 1
      End If

   Loop Until iLBound > iUBound

   IndexedBinarySearchAny = -1
End Function

Public Function IndexedBinarySearchLong(ByRef iArray() As Long, _
                                        ByRef iIndexArray() As Long, _
                                        ByVal iFind As Long) As Long
   Dim iLBound     As Long
   Dim iUBound     As Long
   Dim iMiddle     As Long
   Dim bInitialize As Boolean
   iLBound = LBound(iArray)
   iUBound = UBound(iArray)
   'initialize the index array if necessary
   bInitialize = False

   If UBound(iIndexArray) = -1 Then bInitialize = True Else If iIndexArray(LBound(iIndexArray)) = 0 And iIndexArray(UBound(iIndexArray)) = 0 Then bInitialize = True
   If bInitialize = True Then CreateIndex iIndexArray, iArray

   Do
      iMiddle = (iLBound + iUBound) \ 2

      If iArray(iIndexArray(iMiddle)) = iFind Then
         IndexedBinarySearchLong = iIndexArray(iMiddle)
         Exit Function

      ElseIf iArray(iIndexArray(iMiddle)) < iFind Then
         iLBound = iMiddle + 1
      Else
         iUBound = iMiddle - 1
      End If

   Loop Until iLBound > iUBound

   IndexedBinarySearchLong = -1
End Function

Public Function IndexedBinarySearchString(ByRef sArray() As String, _
                                          ByRef iIndexArray() As Long, _
                                          ByVal sFind As String) As Long
   Dim iLBound     As Long
   Dim iUBound     As Long
   Dim iMiddle     As Long
   Dim bInitialize As Boolean
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)
   'initialize the index array if necessary
   bInitialize = False

   If UBound(iIndexArray) = -1 Then bInitialize = True Else If iIndexArray(LBound(iIndexArray)) = 0 And iIndexArray(UBound(iIndexArray)) = 0 Then bInitialize = True
   If bInitialize = True Then CreateIndex iIndexArray, sArray

   Do
      iMiddle = (iLBound + iUBound) \ 2

      If sArray(iIndexArray(iMiddle)) = sFind Then
         IndexedBinarySearchString = iIndexArray(iMiddle)
         Exit Function

      ElseIf sArray(iIndexArray(iMiddle)) < sFind Then
         iLBound = iMiddle + 1
      Else
         iUBound = iMiddle - 1
      End If

   Loop Until iLBound > iUBound

   IndexedBinarySearchString = -1
End Function

Public Function SequentialSearchAnyArray(ByRef vArray As Variant, _
                                         ByVal vFind As Variant) As Long
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long

   If Not IsArray(vArray) Then Exit Function
   iLBound = LBound(vArray)
   iUBound = UBound(vArray)

   For i = iLBound To iUBound

      If vArray(i) = vFind Then SequentialSearchAnyArray = i: Exit Function
   Next i

   SequentialSearchAnyArray = -1
End Function

Public Function SequentialSearchLongArray(ByRef iArray() As Long, _
                                          ByVal iFind As Long) As Long
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long
   iLBound = LBound(iArray)
   iUBound = UBound(iArray)

   For i = iLBound To iUBound

      If iArray(i) = iFind Then SequentialSearchLongArray = i: Exit Function
   Next i

   SequentialSearchLongArray = -1
End Function

Public Function SequentialSearchStringArray(ByRef sArray() As String, _
                                            ByVal sFind As String) As Long
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)

   For i = iLBound To iUBound

      If sArray(i) = sFind Then SequentialSearchStringArray = i: Exit Function
   Next i

   SequentialSearchStringArray = -1
End Function

Public Function isInAnyArray(ByRef vArray As Variant, _
                             ByVal vFind As Variant) As Boolean

   If Not IsArray(vArray) Then isInAnyArray = False: Exit Function
   isInAnyArray = IIf(SequentialSearchAnyArray(vArray, vFind) = -1, False, True)
End Function

Public Function isInLongArray(ByRef iArray() As Long, _
                              ByVal iFind As Long) As Boolean
   isInLongArray = IIf(SequentialSearchLongArray(iArray, iFind) = -1, False, True)
End Function

Public Function isInStringArray(ByRef sArray() As String, _
                                ByVal sFind As String) As Boolean
   isInStringArray = IIf(SequentialSearchStringArray(sArray, sFind) = -1, False, True)
End Function

' //////////
' // Sort //
' //////////
Public Sub ShellSortAny(ByRef vArray As Variant, _
                        Optional ByVal SortOrder As SortOrder = SortAscending)
   Dim i          As Long   ' Loop Counter
   Dim j          As Long
   Dim iLBound    As Long
   Dim iUBound    As Long
   Dim iMax       As Long
   Dim vTemp      As Variant
   Dim distance   As Long
   Dim bSortOrder As Boolean

   If Not IsArray(vArray) Then Exit Sub
   iLBound = LBound(vArray)
   iUBound = UBound(vArray)
   bSortOrder = IIf(SortOrder = SortAscending, False, True)
   iMax = iUBound - iLBound + 1

   Do
      distance = distance * 3 + 1
   Loop Until distance > iMax

   Do
      distance = distance \ 3

      For i = distance + iLBound To iUBound
         vTemp = vArray(i)
         j = i

         Do While (vArray(j - distance) > vTemp) Xor bSortOrder
            vArray(j) = vArray(j - distance)
            j = j - distance

            If j - distance < iLBound Then Exit Do
         Loop

         vArray(j) = vTemp
      Next i

   Loop Until distance = 1

End Sub

Public Sub ShellSortLong(ByRef iArray() As Long, _
                         Optional ByVal SortOrder As SortOrder = SortAscending)
   Dim i          As Long   ' Loop Counter
   Dim j          As Long
   Dim iLBound    As Long
   Dim iUBound    As Long
   Dim iMax       As Long
   Dim iTemp      As Long
   Dim distance   As Long
   Dim bSortOrder As Boolean
   iLBound = LBound(iArray)
   iUBound = UBound(iArray)
   bSortOrder = IIf(SortOrder = SortAscending, False, True)
   iMax = iUBound - iLBound + 1

   Do
      distance = distance * 3 + 1
   Loop Until distance > iMax

   Do
      distance = distance \ 3

      For i = distance + iLBound To iUBound
         iTemp = iArray(i)
         j = i

         Do While (iArray(j - distance) > iTemp) Xor bSortOrder
            iArray(j) = iArray(j - distance)
            j = j - distance

            If j - distance < iLBound Then Exit Do
         Loop

         iArray(j) = iTemp
      Next i

   Loop Until distance = 1

End Sub

Public Sub ShellSortString(ByRef sArray() As String, _
                           Optional ByVal SortOrder As SortOrder = SortAscending)
   Dim i          As Long   ' Loop Counter
   Dim j          As Long
   Dim iLBound    As Long
   Dim iUBound    As Long
   Dim iMax       As Long
   Dim sTemp      As String
   Dim distance   As Long
   Dim bSortOrder As Boolean
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)
   bSortOrder = IIf(SortOrder = SortAscending, False, True)
   iMax = iUBound - iLBound + 1

   Do
      distance = distance * 3 + 1
   Loop Until distance > iMax

   Do
      distance = distance \ 3

      For i = distance + iLBound To iUBound
         CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(i)), 4 'sTemp = sArray(i)
         j = i

         Do While (sArray(j - distance) > sTemp) Xor bSortOrder
            CopyMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sArray(j - distance)), 4 'sArray(j) = sArray(j - distance)
            j = j - distance

            If j - distance < iLBound Then Exit Do
         Loop

         CopyMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sTemp), 4 'sArray(j) = sTemp
      Next i

   Loop Until distance = 1

   ' delete temp var (sTemp)
   i = 0
   CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4
End Sub

Public Sub TriQuickSortAny(ByRef vArray As Variant, _
                           Optional ByVal SortOrder As SortOrder = SortAscending)
   Dim iLBound As Long
   Dim iUBound As Long
   Dim i       As Long
   Dim j       As Long
   Dim vTemp   As Variant

   If Not IsArray(vArray) Then Exit Sub
   iLBound = LBound(vArray)
   iUBound = UBound(vArray)
   ' *NOTE*  the value 4 is VERY important here !!!
   ' DO NOT CHANGE 4 FOR A LOWER VALUE !!!
   TriQuickSortAny2 vArray, 4, iLBound, iUBound
   InsertionSortAny vArray, iLBound, iUBound

   If SortOrder = SortDescending Then ReverseAnyArray vArray
End Sub

Public Sub TriQuickSortLong(ByRef iArray() As Long, _
                            Optional ByVal SortOrder As SortOrder = SortAscending)
   Dim iLBound As Long
   Dim iUBound As Long
   Dim i       As Long
   Dim j       As Long
   Dim iTemp   As Long
   iLBound = LBound(iArray)
   iUBound = UBound(iArray)
   ' *NOTE*  the value 4 is VERY important here !!!
   ' DO NOT CHANGE 4 FOR A LOWER VALUE !!!
   TriQuickSortLong2 iArray, 4, iLBound, iUBound
   InsertionSortLong iArray, iLBound, iUBound

   If SortOrder = SortDescending Then ReverseLongArray iArray
End Sub

Public Sub TriQuickSortString(ByRef sArray() As String, _
                              Optional ByVal SortOrder As SortOrder = SortAscending)
   Dim iLBound As Long
   Dim iUBound As Long
   Dim i       As Long
   Dim j       As Long
   Dim sTemp   As String
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)
   ' *NOTE*  the value 4 is VERY important here !!!
   ' DO NOT CHANGE 4 FOR A LOWER VALUE !!!
   TriQuickSortString2 sArray, 4, iLBound, iUBound
   InsertionSortString sArray, iLBound, iUBound

   If SortOrder = SortDescending Then ReverseStringArray sArray
End Sub

Public Sub IndexedShellSortAny(ByRef vArray As Variant, _
                               ByRef iIndexArray() As Long, _
                               Optional ByVal SortOrder As SortOrder = SortAscending)
   Dim i           As Long   ' Loop Counter
   Dim j           As Long
   Dim iLBound     As Long
   Dim iUBound     As Long
   Dim iMax        As Long
   Dim vTemp       As Variant
   Dim iIndexTemp  As Long
   Dim distance    As Long
   Dim bInitialize As Boolean
   Dim bSortOrder  As Boolean

   If Not IsArray(vArray) Then Exit Sub
   iLBound = LBound(vArray)
   iUBound = UBound(vArray)
   'initialize the index array if necessary
   bInitialize = False

   If UBound(iIndexArray) = -1 Then bInitialize = True Else If iIndexArray(LBound(iIndexArray)) = 0 And iIndexArray(UBound(iIndexArray)) = 0 Then bInitialize = True
   If bInitialize = True Then CreateIndex iIndexArray, vArray
   bSortOrder = IIf(SortOrder = SortAscending, False, True)
   iMax = iUBound - iLBound + 1

   Do
      distance = distance * 3 + 1
   Loop Until distance > iMax

   Do
      distance = distance \ 3

      For i = distance + iLBound To iUBound
         iIndexTemp = iIndexArray(i)
         vTemp = vArray(iIndexTemp)
         j = i

         Do While (vArray(iIndexArray(j - distance)) > vTemp) Xor bSortOrder
            iIndexArray(j) = iIndexArray(j - distance)
            j = j - distance

            If j - distance < iLBound Then Exit Do
         Loop

         iIndexArray(j) = iIndexTemp
      Next i

   Loop Until distance = 1

End Sub

Public Sub IndexedShellSortLong(ByRef iArray() As Long, _
                                ByRef iIndexArray() As Long, _
                                Optional ByVal SortOrder As SortOrder = SortAscending)
   Dim i           As Long   ' Loop Counter
   Dim j           As Long
   Dim iLBound     As Long
   Dim iUBound     As Long
   Dim iMax        As Long
   Dim iTemp       As Long
   Dim iIndexTemp  As Long
   Dim distance    As Long
   Dim bInitialize As Boolean
   Dim bSortOrder  As Boolean
   iLBound = LBound(iArray)
   iUBound = UBound(iArray)
   'initialize the index array if necessary
   bInitialize = False

   If UBound(iIndexArray) = -1 Then bInitialize = True Else If iIndexArray(LBound(iIndexArray)) = 0 And iIndexArray(UBound(iIndexArray)) = 0 Then bInitialize = True
   If bInitialize = True Then CreateIndex iIndexArray, iArray
   bSortOrder = IIf(SortOrder = SortAscending, False, True)
   iMax = iUBound - iLBound + 1

   Do
      distance = distance * 3 + 1
   Loop Until distance > iMax

   Do
      distance = distance \ 3

      For i = distance + iLBound To iUBound
         iIndexTemp = iIndexArray(i)
         iTemp = iArray(iIndexTemp)
         j = i

         Do While (iArray(iIndexArray(j - distance)) > iTemp) Xor bSortOrder
            iIndexArray(j) = iIndexArray(j - distance)
            j = j - distance

            If j - distance < iLBound Then Exit Do
         Loop

         iIndexArray(j) = iIndexTemp
      Next i

   Loop Until distance = 1

End Sub

Public Sub IndexedShellSortString(ByRef sArray() As String, _
                                  ByRef iIndexArray() As Long, _
                                  Optional ByVal SortOrder As SortOrder = SortAscending)
   Dim i           As Long   ' Loop Counter
   Dim j           As Long
   Dim iLBound     As Long
   Dim iUBound     As Long
   Dim iMax        As Long
   Dim sTemp       As String
   Dim iIndexTemp  As Long
   Dim distance    As Long
   Dim bInitialize As Boolean
   Dim bSortOrder  As Boolean
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)
   'initialize the index array if necessary
   bInitialize = False

   If UBound(iIndexArray) = -1 Then bInitialize = True Else If iIndexArray(LBound(iIndexArray)) = 0 And iIndexArray(UBound(iIndexArray)) = 0 Then bInitialize = True
   If bInitialize = True Then CreateIndex iIndexArray, sArray
   bSortOrder = IIf(SortOrder = SortAscending, False, True)
   iMax = iUBound - iLBound + 1

   Do
      distance = distance * 3 + 1
   Loop Until distance > iMax

   Do
      distance = distance \ 3

      For i = distance + iLBound To iUBound
         iIndexTemp = iIndexArray(i)
         sTemp = sArray(iIndexTemp)
         j = i

         Do While (sArray(iIndexArray(j - distance)) > sTemp) Xor bSortOrder
            iIndexArray(j) = iIndexArray(j - distance)
            j = j - distance

            If j - distance < iLBound Then Exit Do
         Loop

         iIndexArray(j) = iIndexTemp
      Next i

   Loop Until distance = 1

End Sub

Public Sub IndexedTriQuickSortAny(ByRef vArray As Variant, _
                                  ByRef iIndexArray() As Long, _
                                  Optional ByVal SortOrder As SortOrder = SortAscending)
   Dim iLBound     As Long
   Dim iUBound     As Long
   Dim bInitialize As Boolean

   If Not IsArray(vArray) Then Exit Sub
   'initialize the index array if necessary
   bInitialize = False

   If UBound(iIndexArray) = -1 Then bInitialize = True Else If iIndexArray(LBound(iIndexArray)) = 0 And iIndexArray(UBound(iIndexArray)) = 0 Then bInitialize = True
   If bInitialize = True Then CreateIndex iIndexArray, vArray
   iLBound = LBound(vArray)
   iUBound = UBound(vArray)
   ' *NOTE*  the value 4 is VERY important here !!!
   ' DO NOT CHANGE 4 FOR A LOWER VALUE !!!
   IndexedTriQuickSortAny2 vArray, iIndexArray, 4, iLBound, iUBound
   IndexedInsertionSortAny vArray, iIndexArray, iLBound, iUBound

   If SortOrder = SortDescending Then ReverseLongArray iIndexArray
End Sub

Public Sub IndexedTriQuickSortLong(ByRef iArray() As Long, _
                                   ByRef iIndexArray() As Long, _
                                   Optional ByVal SortOrder As SortOrder = SortAscending)
   Dim iLBound     As Long
   Dim iUBound     As Long
   Dim i           As Long
   Dim j           As Long
   Dim bInitialize As Boolean

   If Not IsArray(iArray) Then Exit Sub
   'initialize the index array if necessary
   bInitialize = False

   If UBound(iIndexArray) = -1 Then bInitialize = True Else If iIndexArray(LBound(iIndexArray)) = 0 And iIndexArray(UBound(iIndexArray)) = 0 Then bInitialize = True
   If bInitialize = True Then CreateIndex iIndexArray, iArray
   iLBound = LBound(iArray)
   iUBound = UBound(iArray)
   ' *NOTE*  the value 4 is VERY important here !!!
   ' DO NOT CHANGE 4 FOR A LOWER VALUE !!!
   IndexedTriQuickSortLong2 iArray, iIndexArray, 4, iLBound, iUBound
   IndexedInsertionSortLong iArray, iIndexArray, iLBound, iUBound

   If SortOrder = SortDescending Then ReverseLongArray iIndexArray
End Sub

Public Sub IndexedTriQuickSortString(ByRef sArray() As String, _
                                     ByRef iIndexArray() As Long, _
                                     Optional ByVal SortOrder As SortOrder = SortAscending)
   Dim iLBound     As Long
   Dim iUBound     As Long
   Dim i           As Long
   Dim j           As Long
   Dim iPos        As Long
   Dim bInitialize As Boolean

   If Not IsArray(sArray) Then Exit Sub
   'initialize the index array if necessary
   bInitialize = False

   If UBound(iIndexArray) = -1 Then bInitialize = True Else If iIndexArray(LBound(iIndexArray)) = 0 And iIndexArray(UBound(iIndexArray)) = 0 Then bInitialize = True
   If bInitialize = True Then CreateIndex iIndexArray, sArray
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)
   ' *NOTE*  the value 4 is VERY important here !!!
   ' DO NOT CHANGE 4 FOR A LOWER VALUE !!!
   IndexedTriQuickSortString2 sArray, iIndexArray, 4, iLBound, iUBound
   IndexedInsertionSortString sArray, iIndexArray, iLBound, iUBound

   If SortOrder = SortDescending Then ReverseLongArray iIndexArray
End Sub

Public Function isSortedAnyArray(ByRef vArray As Variant, _
                                 Optional ByVal SortOrder As SortOrder = SortAscending) As Boolean
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iStep   As Long

   If Not IsArray(vArray) Then isSortedAnyArray = False: Exit Function
   iStep = IIf(SortOrder = SortAscending, 1, -1)
   iLBound = IIf(SortOrder = SortAscending, LBound(vArray), UBound(vArray))
   iUBound = IIf(SortOrder = SortAscending, UBound(vArray), LBound(vArray)) - iStep

   For i = iLBound To iUBound Step iStep

      If vArray(i) > vArray(i + iStep) Then isSortedAnyArray = False: Exit Function
   Next i

   isSortedAnyArray = True
End Function

Public Function isSortedLongArray(ByRef iArray() As Long, _
                                  Optional ByVal SortOrder As SortOrder = SortAscending) As Boolean
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iStep   As Long
   iStep = IIf(SortOrder = SortAscending, 1, -1)
   iLBound = IIf(SortOrder = SortAscending, LBound(iArray), UBound(iArray))
   iUBound = IIf(SortOrder = SortAscending, UBound(iArray), LBound(iArray)) - iStep

   For i = iLBound To iUBound Step iStep

      If iArray(i) > iArray(i + iStep) Then isSortedLongArray = False: Exit Function
   Next i

   isSortedLongArray = True
End Function

Public Function isSortedStringArray(ByRef sArray() As String, _
                                    Optional ByVal SortOrder As SortOrder = SortAscending) As Boolean
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iStep   As Long
   iStep = IIf(SortOrder = SortAscending, 1, -1)
   iLBound = IIf(SortOrder = SortAscending, LBound(sArray), UBound(sArray))
   iUBound = IIf(SortOrder = SortAscending, UBound(sArray), LBound(sArray)) - iStep

   For i = iLBound To iUBound Step iStep

      If sArray(i) > sArray(i + iStep) Then isSortedStringArray = False: Exit Function
   Next i

   isSortedStringArray = True
End Function

Public Function isSortedIndexedAnyArray(ByRef vArray As Variant, _
                                        ByRef iIndexArray() As Long, _
                                        Optional ByVal SortOrder As SortOrder = SortAscending) As Boolean
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iStep   As Long

   If Not IsArray(vArray) Then isSortedIndexedAnyArray = False: Exit Function
   iStep = IIf(SortOrder = SortAscending, 1, -1)
   iLBound = IIf(SortOrder = SortAscending, LBound(vArray), UBound(vArray))
   iUBound = IIf(SortOrder = SortAscending, UBound(vArray), LBound(vArray)) - iStep

   For i = iLBound To iUBound Step iStep

      If vArray(iIndexArray(i)) > vArray(iIndexArray(i + iStep)) Then isSortedIndexedAnyArray = False: Exit Function
   Next i

   isSortedIndexedAnyArray = True
End Function

Public Function isSortedIndexedLongArray(ByRef iArray() As Long, _
                                         ByRef iIndexArray() As Long, _
                                         Optional ByVal SortOrder As SortOrder = SortAscending) As Boolean
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iStep   As Long
   iStep = IIf(SortOrder = SortAscending, 1, -1)
   iLBound = IIf(SortOrder = SortAscending, LBound(iArray), UBound(iArray))
   iUBound = IIf(SortOrder = SortAscending, UBound(iArray), LBound(iArray)) - iStep

   For i = iLBound To iUBound Step iStep

      If iArray(iIndexArray(i)) > iArray(iIndexArray(i + iStep)) Then isSortedIndexedLongArray = False: Exit Function
   Next i

   isSortedIndexedLongArray = True
End Function

Public Function isSortedIndexedStringArray(ByRef sArray() As String, _
                                           ByRef iIndexArray() As Long, _
                                           Optional ByVal SortOrder As SortOrder = SortAscending) As Boolean
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long
   Dim iStep   As Long
   iStep = IIf(SortOrder = SortAscending, 1, -1)
   iLBound = IIf(SortOrder = SortAscending, LBound(sArray), UBound(sArray))
   iUBound = IIf(SortOrder = SortAscending, UBound(sArray), LBound(sArray)) - iStep

   For i = iLBound To iUBound Step iStep

      If sArray(iIndexArray(i)) > sArray(iIndexArray(i + iStep)) Then isSortedIndexedStringArray = False: Exit Function
   Next i

   isSortedIndexedStringArray = True
End Function

' /////////////////////
' // Synchronisation //
' /////////////////////
Public Sub SynchroniseIndexedAnyArray(ByRef vArray As Variant, _
                                      ByRef iIndexArray() As Long)
   Dim i            As Long
   Dim iLBound      As Long
   Dim iUBound      As Long
   Dim vArrayTemp() As Variant

   If Not IsArray(vArray) Then Exit Sub
   iLBound = LBound(vArray)
   iUBound = UBound(vArray)
   ' vArrayTemp serves as a copy of vArray so that the synchronise effect is saved directly on vArray.
   CopyAnyArray vArray, vArrayTemp

   For i = iLBound To iUBound
      vArray(i) = vArrayTemp(iIndexArray(i))
   Next i

   ' recreate the index
   CreateIndex iIndexArray, vArray
   Erase vArrayTemp
End Sub

Public Sub SynchroniseIndexedLongArray(ByRef iArray() As Long, _
                                       ByRef iIndexArray() As Long)
   Dim i            As Long
   Dim iLBound      As Long
   Dim iUBound      As Long
   Dim iArrayTemp() As Long
   iLBound = LBound(iArray)
   iUBound = UBound(iArray)
   ' because we want our synchronise effect to be saved directly on iArray.
   MoveLongArray iArray, iArrayTemp
   ReDim iArray(iLBound To iUBound)

   For i = iLBound To iUBound
      iArray(i) = iArrayTemp(iIndexArray(i))
   Next i

   ' recreate the index
   CreateIndex iIndexArray, iArray
   Erase iArrayTemp
End Sub

Public Sub SynchroniseIndexedStringArray(ByRef sArray() As String, _
                                         ByRef iIndexArray() As Long)
   Dim i            As Long
   Dim iLBound      As Long
   Dim iUBound      As Long
   Dim sArrayTemp() As String
   Dim iNullArray() As Long ' we use this array to imitate ZeroMemory behavior using CopyMemory with 0's
   Dim nBytes       As Long
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)
   ReDim iNullArray(iUBound - iLBound + 1)
   nBytes = (iUBound - iLBound + 1) * 4
   ' because we want our synchronise effect to be saved directly on sArray.
   MoveStringArray sArray, sArrayTemp
   ReDim sArray(iLBound To iUBound)

   For i = iLBound To iUBound
      CopyMemory ByVal VarPtr(sArray(i)), ByVal VarPtr(sArrayTemp(iIndexArray(i))), 4
      'sArray(i) = sArrayTemp(iIndexArray(i))
   Next i

   ' *NOTE* for an unexplicable reason, ZeroMemory is far less stable to use than CopyMemory. (incompatible with WinXP)
   'ZeroMemory ByVal VarPtr(sArraySource(iLBound)), nBytes
   CopyMemory ByVal VarPtr(sArrayTemp(iLBound)), ByVal VarPtr(iNullArray(0)), nBytes
   ' recreate the index
   CreateIndex iIndexArray, sArray
   Erase sArrayTemp
End Sub

' ///////////////
' // Copy/Move //
' ///////////////
Public Sub CopyAnyArray(ByRef vArraySource As Variant, _
                        ByRef vArrayDest As Variant)
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long

   If (Not IsArray(vArraySource)) Or (Not IsArray(vArrayDest)) Then Exit Sub
   iLBound = LBound(vArraySource)
   iUBound = UBound(vArraySource)
   ReDim vArrayDest(iLBound To iUBound)

   For i = iLBound To iUBound
      vArrayDest(i) = vArraySource(i)
   Next i

End Sub

Public Sub CopyLongArray(ByRef iArraySource() As Long, _
                         ByRef iArrayDest() As Long)
   ReDim iArrayDest(LBound(iArraySource) To UBound(iArraySource))
   CopyMemory iArrayDest(0), iArraySource(0), (UBound(iArraySource) - LBound(iArraySource) + 1) * Len(iArraySource(0))
End Sub

Public Sub CopyStringArray(ByRef sArraySource() As String, _
                           ByRef sArrayDest() As String)
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long
   iLBound = LBound(sArraySource)
   iUBound = UBound(sArraySource)
   ReDim sArrayDest(iLBound To iUBound)

   For i = iLBound To iUBound
      sArrayDest(i) = sArraySource(i)  ' cannot CopyMemory !
   Next i

End Sub

Public Sub MoveAnyArray(ByRef vArraySource As Variant, _
                        ByRef vArrayDest As Variant)

   If (Not IsArray(vArraySource)) Or (Not IsArray(vArrayDest)) Then Exit Sub
   CopyAnyArray vArraySource, vArrayDest
   Erase vArraySource
End Sub

Public Sub MoveLongArray(ByRef iArraySource() As Long, _
                         ByRef iArrayDest() As Long)
   CopyLongArray iArraySource, iArrayDest
   Erase iArraySource
End Sub

Public Sub MoveStringArray(ByRef sArraySource() As String, _
                           ByRef sArrayDest() As String)
   Dim iLBound      As Long
   Dim iUBound      As Long
   Dim nBytes       As Long
   Dim iNullArray() As Long ' we use this array to imitate ZeroMemory behavior using CopyMemory with 0's
   iLBound = LBound(sArraySource)
   iUBound = UBound(sArraySource)
   ReDim iNullArray(iUBound - iLBound + 1)
   nBytes = (iUBound - iLBound + 1) * 4
   ReDim sArrayDest(iLBound To iUBound) As String
   CopyMemory ByVal VarPtr(sArrayDest(iLBound)), ByVal VarPtr(sArraySource(iLBound)), nBytes
   ' *NOTE* for an unexplicable reason, ZeroMemory is far less stable to use than CopyMemory. (incompatible with WinXP)
   'ZeroMemory ByVal VarPtr(sArraySource(iLBound)), nBytes
   CopyMemory ByVal VarPtr(sArraySource(iLBound)), ByVal VarPtr(iNullArray(0)), nBytes
   Erase sArraySource
End Sub

Public Sub MergeAnyArray(ByRef vArraySource As Variant, _
                         ByRef vArrayDest As Variant, _
                         Optional ByVal iPos As Long = -1)
   Dim i        As Long
   Dim iLBound  As Long
   Dim iUBound  As Long
   Dim iUBound2 As Long
   Dim iTemp    As Long

   If (Not IsArray(vArraySource)) Or (Not IsArray(vArrayDest)) Then Exit Sub
   iLBound = LBound(vArraySource)
   iUBound = UBound(vArraySource)
   iUBound2 = UBound(vArrayDest)
   iTemp = iUBound - iLBound + 1

   If (iPos > UBound(vArrayDest) + 1) Or (iPos = -1) Then iPos = UBound(vArrayDest) + 1
   If iPos < 0 Then iPos = 0
   ReDim Preserve vArrayDest(LBound(vArrayDest) To UBound(vArrayDest) + iTemp)

   For i = iUBound2 To iPos Step -1
      vArrayDest(i + iTemp) = vArrayDest(i)
   Next i

   iUBound = iPos + iTemp - 1

   For i = iPos To iUBound
      vArrayDest(i) = vArraySource(i - iPos)
   Next i

   Erase vArraySource
End Sub

Public Sub MergeLongArray(ByRef iArraySource() As Long, _
                          ByRef iArrayDest() As Long, _
                          Optional ByVal iPos As Long = -1)
   Dim i        As Long
   Dim iLBound  As Long
   Dim iUBound  As Long
   Dim iUBound2 As Long
   Dim iTemp    As Long
   iLBound = LBound(iArraySource)
   iUBound = UBound(iArraySource)
   iUBound2 = UBound(iArrayDest)
   iTemp = iUBound - iLBound + 1
   ReDim Preserve iArrayDest(LBound(iArrayDest) To iUBound2 + iTemp)

   If (iPos > iUBound2 + 1) Or (iPos = -1) Then
      iPos = iUBound2 + 1
   Else

      If iPos < 0 Then iPos = 0
      CopyMemory iArrayDest(iPos + iTemp), iArrayDest(iPos), (iUBound2 - LBound(iArrayDest) - iPos + 1) * Len(iArrayDest(iPos))
   End If

   CopyMemory iArrayDest(iPos), iArraySource(0), iTemp * Len(iArrayDest(iPos))
   Erase iArraySource
End Sub

Public Sub MergeStringArray(ByRef sArraySource() As String, _
                            ByRef sArrayDest() As String, _
                            Optional ByVal iPos As Long = -1)
   Dim i            As Long
   Dim iLBound      As Long
   Dim iUBound      As Long
   Dim iUBound2     As Long
   Dim iTemp        As Long
   Dim iNull        As Long
   Dim iNullArray() As Long ' we use this array to imitate ZeroMemory behavior using CopyMemory with 0's
   iLBound = LBound(sArraySource)
   iUBound = UBound(sArraySource)
   iUBound2 = UBound(sArrayDest)
   iTemp = iUBound - iLBound + 1
   ReDim iNullArray(iTemp)
   ReDim Preserve sArrayDest(LBound(sArrayDest) To iUBound2 + iTemp)

   If (iPos > iUBound2 + 1) Or (iPos = -1) Then
      iPos = iUBound2 + 1
   Else

      If iPos < 0 Then iPos = 0
      CopyMemory ByVal VarPtr(sArrayDest(iPos + iTemp)), ByVal VarPtr(sArrayDest(iPos)), (iUBound2 - LBound(sArrayDest) - iPos + 1) * 4
   End If

   iTemp = iTemp * 4
   CopyMemory ByVal VarPtr(sArrayDest(iPos)), ByVal VarPtr(sArraySource(iLBound)), iTemp
   ' *NOTE* for an unexplicable reason, ZeroMemory is far less stable to use than CopyMemory. (incompatible with WinXP)
   'ZeroMemory ByVal VarPtr(sArraySource(iLBound)), iTemp * 4
   CopyMemory ByVal VarPtr(sArraySource(iLBound)), ByVal VarPtr(iNullArray(0)), iTemp
   Erase sArraySource
End Sub

' ///////////////
' // Save/Load //
' ///////////////
Public Function SaveLongArray(ByRef iArray() As Long) As String
   Dim iLBound  As Long
   Dim iUBound  As Long
   Dim iUBound2 As Long
   Dim i        As Long
   Dim s()      As Byte
   iLBound = LBound(iArray)
   iUBound = UBound(iArray)
   iUBound2 = 3
   ReDim s(iUBound2)
   CopyMemory ByVal VarPtr(s(0)), iUBound - iLBound + 1, 4 ' number of element

   For i = iLBound To iUBound
      iUBound2 = iUBound2 + 4
      ReDim Preserve s(iUBound2)
      CopyMemory ByVal VarPtr(s(iUBound2 - 3)), iArray(i), 4
   Next i

   ' SaveLongArray = s  ' this does not works (!?!)
   SaveLongArray = Space(iUBound2 + 1)

   For i = 0 To iUBound2
      Mid(SaveLongArray, i + 1, 1) = Chr(s(i))
   Next i

End Function

Public Function SaveStringArray(ByRef sArray() As String) As String
   Dim iLBound  As Long
   Dim iUBound  As Long
   Dim iUBound2 As Long
   Dim i        As Long
   Dim iLen     As Long
   Dim s()      As Byte
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)
   iUBound2 = 3
   iLen = iUBound - iLBound + 1
   ReDim s(iUBound2)
   CopyMemory ByVal VarPtr(s(0)), iLen, 4 ' number of element

   For i = iLBound To iUBound
      iLen = Len(sArray(i))
      iUBound2 = iUBound2 + iLen + 4
      ReDim Preserve s(iUBound2)
      CopyMemory ByVal VarPtr(s(iUBound2 - iLen - 3)), iLen, 4 ' length of nth element
      CopyMemory ByVal VarPtr(s(iUBound2 - iLen + 1)), ByVal sArray(i), iLen ' data
   Next i

   ' SaveStringArray = s  ' this does not works (!?!)
   SaveStringArray = Space(iUBound2 + 1)

   For i = 0 To iUBound2
      Mid(SaveStringArray, i + 1, 1) = Chr(s(i))
   Next i

End Function

Public Sub LoadLongArray(ByRef iArray() As Long, _
                         ByRef sString As String)
   Dim iUBound As Long
   Dim i       As Long
   Dim iPos    As Long
   Dim s()     As Byte

   If Len(sString) = 0 Then Exit Sub
   ' we copy the string to a byte array to avoid unicode bugs (strings CAN be saved in unicode in memory)
   ReDim s(Len(sString) - 1)
   CopyMemory ByVal VarPtr(s(0)), ByVal sString, Len(sString)
   CopyMemory ByVal VarPtr(iUBound), ByVal VarPtr(s(0)), 4 ' number of elements
   iUBound = iUBound - 1
   ReDim iArray(iUBound)
   iPos = 0

   For i = 0 To iUBound
      iPos = iPos + 4
      CopyMemory ByVal VarPtr(iArray(i)), ByVal VarPtr(s(iPos)), 4
   Next i

End Sub

Public Sub LoadStringArray(ByRef sArray() As String, _
                           ByRef sString As String)
   Dim iUBound As Long
   Dim i       As Long
   Dim iPos    As Long
   Dim iLen    As Long
   Dim s()     As Byte

   If Len(sString) = 0 Then Exit Sub
   ' we copy the string to a byte array to avoid unicode bugs (strings CAN be saved in unicode in memory)
   ReDim s(Len(sString) - 1)
   CopyMemory ByVal VarPtr(s(0)), ByVal sString, Len(sString)
   CopyMemory ByVal VarPtr(iUBound), ByVal VarPtr(s(0)), 4 ' number of elements
   iUBound = iUBound - 1
   ReDim sArray(iUBound)
   iPos = 0

   For i = 0 To iUBound
      iPos = iPos + 4
      CopyMemory ByVal VarPtr(iLen), ByVal VarPtr(s(iPos)), 4 ' length of string

      If iLen > 0 Then
         sArray(i) = Mid(sString, iPos + 5, iLen)
         iPos = iPos + iLen
      Else
         sArray(i) = vbNullString
      End If

   Next i

End Sub

' ////////////
' // Others //
' ////////////
' Returns an array of the type of the first sent argument.
Public Function CreateArray(ParamArray values() As Variant) As Variant
   Dim i       As Long
   Dim iUBound As Long
   Dim vTemp   As Variant
   iUBound = UBound(values)

   ' we can't use the vbObject constant for objects because the VarType() function might return the type of the object's default property
   If IsObject(values(0)) Then
      ReDim oObjectArray(0 To iUBound) As Object

      For i = 0 To iUBound
         Set oObjectArray(i) = values(i)
      Next i

      CreateArray = oObjectArray()
      Exit Function

   End If

   Select Case VarType(values(0))

      Case vbLong
         ReDim lArray(0 To iUBound) As Long
         vTemp = lArray()

      Case vbString
         ReDim sArray(0 To iUBound) As String
         vTemp = sArray()

      Case vbInteger
         ReDim iArray(0 To iUBound) As Integer
         vTemp = iArray()

      Case vbSingle
         ReDim sngArray(0 To iUBound) As Single
         vTemp = sngArray()

      Case vbDouble
         ReDim dArray(0 To iUBound) As Double
         vTemp = dArray()

      Case vbCurrency
         ReDim cArray(0 To iUBound) As Currency
         vTemp = cArray()

      Case vbDate
         ReDim datArray(0 To iUBound) As Date
         vTemp = datArray()

      Case vbBoolean
         ReDim bArray(0 To iUBound) As Boolean
         vTemp = bArray()

      Case Else
         ' unsupported data type (UDT or array)
   End Select

   For i = 0 To iUBound
      vTemp(i) = values(i)
   Next i

   CreateArray = vTemp
End Function

' MsgBox an array. Use for debugging.
Public Sub DebugDumpArray(ByRef vArray As Variant, _
                          Optional ByVal iColumnWidth As Long = 4)
   Dim iLBound As Long
   Dim iUBound As Long
   Dim i       As Long
   Dim j       As Long
   Dim iPos    As Long
   Dim sString As String

   If Not IsArray(vArray) Then Exit Sub
   If iColumnWidth < 1 Then iColumnWidth = 1
   iLBound = LBound(vArray)
   iUBound = UBound(vArray)
   iPos = iLBound - 1
   sString = "Dumping array:" & vbTab & "Type -> " & TypeName(vArray) & " <" & iLBound & " To " & iUBound & ">" & vbCrLf & vbCrLf

   If iUBound > 100 Then iUBound = 100 ' MsgBox can't show over 100 anyway.

   For i = iLBound To iUBound

      If iPos + iColumnWidth > iUBound Then iColumnWidth = iUBound - iPos

      For j = 1 To iColumnWidth
         iPos = iPos + 1
         sString = sString & iPos & ":  " & vArray(iPos) & vbTab
      Next j

      sString = sString & vbCrLf
   Next i

   MsgBox sString
End Sub

Public Sub ReverseAnyArray(ByRef vArray As Variant)
   Dim iLBound As Long
   Dim iUBound As Long

   If Not IsArray(vArray) Then Exit Sub
   iLBound = LBound(vArray)
   iUBound = UBound(vArray)
   While iLBound < iUBound
      SwapAny vArray(iLBound), vArray(iUBound)
      iLBound = iLBound + 1
      iUBound = iUBound - 1
   Wend
End Sub

Public Sub ReverseLongArray(ByRef iArray() As Long)
   Dim iLBound As Long
   Dim iUBound As Long
   iLBound = LBound(iArray)
   iUBound = UBound(iArray)
   While iLBound < iUBound
      SwapLongs iArray(iLBound), iArray(iUBound)
      iLBound = iLBound + 1
      iUBound = iUBound - 1
   Wend
End Sub

Public Sub ReverseStringArray(ByRef sArray() As String)
   Dim iLBound As Long
   Dim iUBound As Long
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)
   While iLBound < iUBound
      SwapStrings sArray(iLBound), sArray(iUBound)
      iLBound = iLBound + 1
      iUBound = iUBound - 1
   Wend
End Sub

' //////////////////
' // Private Subs //
' //////////////////
' this sub is intended for internal usage. It only fills iIndexArray().
Private Sub CreateIndex(ByRef iIndexArray() As Long, _
                        ByRef vSizeArray As Variant)
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long
   iLBound = LBound(vSizeArray)
   iUBound = UBound(vSizeArray)
   ReDim iIndexArray(iLBound To iUBound)

   For i = iLBound To iUBound
      iIndexArray(i) = i
   Next

End Sub

Private Sub TriQuickSortAny2(ByRef vArray As Variant, _
                             ByVal iSplit As Long, _
                             ByVal iMin As Long, _
                             ByVal iMax As Long)
   Dim i     As Long
   Dim j     As Long
   Dim vTemp As Variant

   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.
   If (iMax - iMin) > iSplit Then
      i = (iMax + iMin) / 2

      If vArray(iMin) > vArray(i) Then SwapAny vArray(iMin), vArray(i)
      If vArray(iMin) > vArray(iMax) Then SwapAny vArray(iMin), vArray(iMax)
      If vArray(i) > vArray(iMax) Then SwapAny vArray(i), vArray(iMax)
      j = iMax - 1
      SwapAny vArray(i), vArray(j)
      i = iMin
      vTemp = vArray(j)

      Do
         Do
            i = i + 1
         Loop While vArray(i) < vTemp

         Do
            j = j - 1
         Loop While vArray(j) > vTemp

         If j < i Then Exit Do
         SwapAny vArray(i), vArray(j)
      Loop

      SwapAny vArray(i), vArray(iMax - 1)
      TriQuickSortAny2 vArray, iSplit, iMin, j
      TriQuickSortAny2 vArray, iSplit, i + 1, iMax
   End If

End Sub

Private Sub TriQuickSortLong2(ByRef iArray() As Long, _
                              ByVal iSplit As Long, _
                              ByVal iMin As Long, _
                              ByVal iMax As Long)
   Dim i     As Long
   Dim j     As Long
   Dim iTemp As Long

   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.
   If (iMax - iMin) > iSplit Then
      i = (iMax + iMin) / 2

      If iArray(iMin) > iArray(i) Then SwapLongs iArray(iMin), iArray(i)
      If iArray(iMin) > iArray(iMax) Then SwapLongs iArray(iMin), iArray(iMax)
      If iArray(i) > iArray(iMax) Then SwapLongs iArray(i), iArray(iMax)
      j = iMax - 1
      SwapLongs iArray(i), iArray(j)
      i = iMin
      iTemp = iArray(j)

      Do
         Do
            i = i + 1
         Loop While iArray(i) < iTemp

         Do
            j = j - 1
         Loop While iArray(j) > iTemp

         If j < i Then Exit Do
         SwapLongs iArray(i), iArray(j)
      Loop

      SwapLongs iArray(i), iArray(iMax - 1)
      TriQuickSortLong2 iArray, iSplit, iMin, j
      TriQuickSortLong2 iArray, iSplit, i + 1, iMax
   End If

End Sub

Private Sub TriQuickSortString2(ByRef sArray() As String, _
                                ByVal iSplit As Long, _
                                ByVal iMin As Long, _
                                ByVal iMax As Long)
   Dim i     As Long
   Dim j     As Long
   Dim sTemp As String

   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.
   If (iMax - iMin) > iSplit Then
      i = (iMax + iMin) / 2

      If sArray(iMin) > sArray(i) Then SwapStrings sArray(iMin), sArray(i)
      If sArray(iMin) > sArray(iMax) Then SwapStrings sArray(iMin), sArray(iMax)
      If sArray(i) > sArray(iMax) Then SwapStrings sArray(i), sArray(iMax)
      j = iMax - 1
      SwapStrings sArray(i), sArray(j)
      i = iMin
      CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(j)), 4 ' sTemp = sArray(j)

      Do
         Do
            i = i + 1
         Loop While sArray(i) < sTemp

         Do
            j = j - 1
         Loop While sArray(j) > sTemp

         If j < i Then Exit Do
         SwapStrings sArray(i), sArray(j)
      Loop

      SwapStrings sArray(i), sArray(iMax - 1)
      TriQuickSortString2 sArray, iSplit, iMin, j
      TriQuickSortString2 sArray, iSplit, i + 1, iMax
   End If

   ' clear temp var (sTemp)
   i = 0
   CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4
End Sub

Private Sub IndexedTriQuickSortAny2(ByRef vArray As Variant, _
                                    ByRef iIndexArray() As Long, _
                                    ByVal iSplit As Long, _
                                    ByVal iMin As Long, _
                                    ByVal iMax As Long)
   Dim i     As Long
   Dim j     As Long
   Dim vTemp As Variant

   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.
   If (iMax - iMin) > iSplit Then
      i = (iMax + iMin) / 2

      If vArray(iIndexArray(iMin)) > vArray(iIndexArray(i)) Then SwapLongs iIndexArray(iMin), iIndexArray(i)
      If vArray(iIndexArray(iMin)) > vArray(iIndexArray(iMax)) Then SwapLongs iIndexArray(iMin), iIndexArray(iMax)
      If vArray(iIndexArray(i)) > vArray(iIndexArray(iMax)) Then SwapLongs iIndexArray(i), iIndexArray(iMax)
      j = iMax - 1
      SwapLongs iIndexArray(i), iIndexArray(j)
      i = iMin
      vTemp = vArray(iIndexArray(j))

      Do
         Do
            i = i + 1
         Loop While vArray(iIndexArray(i)) < vTemp

         Do
            j = j - 1
         Loop While vArray(iIndexArray(j)) > vTemp

         If j < i Then Exit Do
         SwapLongs iIndexArray(i), iIndexArray(j)
      Loop

      SwapLongs iIndexArray(i), iIndexArray(iMax - 1)
      IndexedTriQuickSortAny2 vArray, iIndexArray, iSplit, iMin, j
      IndexedTriQuickSortAny2 vArray, iIndexArray, iSplit, i + 1, iMax
   End If

End Sub

Private Sub IndexedTriQuickSortLong2(ByRef iArray() As Long, _
                                     ByRef iIndexArray() As Long, _
                                     ByVal iSplit As Long, _
                                     ByVal iMin As Long, _
                                     ByVal iMax As Long)
   Dim i     As Long
   Dim j     As Long
   Dim iTemp As Long

   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.
   If (iMax - iMin) > iSplit Then
      i = (iMax + iMin) / 2

      If iArray(iIndexArray(iMin)) > iArray(iIndexArray(i)) Then SwapLongs iIndexArray(iMin), iIndexArray(i)
      If iArray(iIndexArray(iMin)) > iArray(iIndexArray(iMax)) Then SwapLongs iIndexArray(iMin), iIndexArray(iMax)
      If iArray(iIndexArray(i)) > iArray(iIndexArray(iMax)) Then SwapLongs iIndexArray(i), iIndexArray(iMax)
      j = iMax - 1
      SwapLongs iIndexArray(i), iIndexArray(j)
      i = iMin
      iTemp = iArray(iIndexArray(j))

      Do
         Do
            i = i + 1
         Loop While iArray(iIndexArray(i)) < iTemp

         Do
            j = j - 1
         Loop While iArray(iIndexArray(j)) > iTemp

         If j < i Then Exit Do
         SwapLongs iIndexArray(i), iIndexArray(j)
      Loop

      SwapLongs iIndexArray(i), iIndexArray(iMax - 1)
      IndexedTriQuickSortLong2 iArray, iIndexArray, iSplit, iMin, j
      IndexedTriQuickSortLong2 iArray, iIndexArray, iSplit, i + 1, iMax
   End If

End Sub

Private Sub IndexedTriQuickSortString2(ByRef sArray() As String, _
                                       ByRef iIndexArray() As Long, _
                                       ByVal iSplit As Long, _
                                       ByVal iMin As Long, _
                                       ByVal iMax As Long)
   Dim i     As Long
   Dim j     As Long
   Dim sTemp As String

   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.
   If (iMax - iMin) > iSplit Then
      i = (iMax + iMin) / 2

      If sArray(iIndexArray(iMin)) > sArray(iIndexArray(i)) Then SwapLongs iIndexArray(iMin), iIndexArray(i)
      If sArray(iIndexArray(iMin)) > sArray(iIndexArray(iMax)) Then SwapLongs iIndexArray(iMin), iIndexArray(iMax)
      If sArray(iIndexArray(i)) > sArray(iIndexArray(iMax)) Then SwapLongs iIndexArray(i), iIndexArray(iMax)
      j = iMax - 1
      SwapLongs iIndexArray(i), iIndexArray(j)
      i = iMin
      CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(iIndexArray(j))), 4 ' sTemp = sArray(iIndexArray(j))

      Do
         Do
            i = i + 1
         Loop While sArray(iIndexArray(i)) < sTemp

         Do
            j = j - 1
         Loop While sArray(iIndexArray(j)) > sTemp

         If j < i Then Exit Do
         SwapLongs iIndexArray(i), iIndexArray(j)
      Loop

      SwapLongs iIndexArray(i), iIndexArray(iMax - 1)
      IndexedTriQuickSortString2 sArray, iIndexArray, iSplit, iMin, j
      IndexedTriQuickSortString2 sArray, iIndexArray, iSplit, i + 1, iMax
   End If

   ' clear temp var (sTemp)
   i = 0
   CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4
End Sub

Private Sub InsertionSortAny(ByRef vArray As Variant, _
                             ByVal iMin As Long, _
                             ByVal iMax As Long)
   Dim i     As Long
   Dim j     As Long
   Dim vTemp As Variant

   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.
   For i = iMin + 1 To iMax
      vTemp = vArray(i)
      j = i

      Do While j > iMin

         If vArray(j - 1) <= vTemp Then Exit Do
         vArray(j) = vArray(j - 1)
         j = j - 1
      Loop

      vArray(j) = vTemp
   Next i

End Sub

Private Sub InsertionSortLong(ByRef iArray() As Long, _
                              ByVal iMin As Long, _
                              ByVal iMax As Long)
   Dim i     As Long
   Dim j     As Long
   Dim iTemp As Long

   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.
   For i = iMin + 1 To iMax
      iTemp = iArray(i)
      j = i

      Do While j > iMin

         If iArray(j - 1) <= iTemp Then Exit Do
         iArray(j) = iArray(j - 1)
         j = j - 1
      Loop

      iArray(j) = iTemp
   Next i

End Sub

Private Sub InsertionSortString(ByRef sArray() As String, _
                                ByVal iMin As Long, _
                                ByVal iMax As Long)
   Dim i     As Long
   Dim j     As Long
   Dim sTemp As String

   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.
   For i = iMin + 1 To iMax
      CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(i)), 4 ' sTemp = sArray(i)
      j = i

      Do While j > iMin

         If sArray(j - 1) <= sTemp Then Exit Do
         CopyMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sArray(j - 1)), 4 ' sArray(j) = sArray(j - 1)
         j = j - 1
      Loop

      CopyMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sTemp), 4
      ' sArray(j) = sTemp
   Next i

   ' clear temp var (sTemp)
   i = 0
   CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4
End Sub

Private Sub IndexedInsertionSortAny(ByRef vArray As Variant, _
                                    ByRef iIndexArray() As Long, _
                                    ByVal iMin As Long, _
                                    ByVal iMax As Long)
   Dim i    As Long
   Dim j    As Long
   Dim iPos As Long

   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.
   For i = iMin + 1 To iMax
      iPos = iIndexArray(i)
      j = i

      Do While j > iMin

         If vArray(iIndexArray(j - 1)) <= vArray(iPos) Then Exit Do
         iIndexArray(j) = iIndexArray(j - 1)
         j = j - 1
      Loop

      iIndexArray(j) = iPos
   Next i

End Sub

Private Sub IndexedInsertionSortLong(ByRef iArray() As Long, _
                                     ByRef iIndexArray() As Long, _
                                     ByVal iMin As Long, _
                                     ByVal iMax As Long)
   Dim i    As Long
   Dim j    As Long
   Dim iPos As Long

   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.
   For i = iMin + 1 To iMax
      iPos = iIndexArray(i)
      j = i

      Do While j > iMin

         If iArray(iIndexArray(j - 1)) <= iArray(iPos) Then Exit Do
         iIndexArray(j) = iIndexArray(j - 1)
         j = j - 1
      Loop

      iIndexArray(j) = iPos
   Next i

End Sub

Private Sub IndexedInsertionSortString(ByRef sArray() As String, _
                                       ByRef iIndexArray() As Long, _
                                       ByVal iMin As Long, _
                                       ByVal iMax As Long)
   Dim i    As Long
   Dim j    As Long
   Dim iPos As Long

   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.
   For i = iMin + 1 To iMax
      iPos = iIndexArray(i)
      j = i

      Do While j > iMin

         If sArray(iIndexArray(j - 1)) <= sArray(iPos) Then Exit Do
         iIndexArray(j) = iIndexArray(j - 1)
         j = j - 1
      Loop

      iIndexArray(j) = iPos
   Next i

End Sub

' Do not redefine if mdlMarton is already loaded
' Ok, I know theses are Private, but I still prefer to do it this way.
#If mdlMarton_Loadedable = False Then
   ' Swaps 2 variants.
   Private Sub SwapAny(ByRef v1 As Variant, _
                       ByRef v2 As Variant)
      Dim V As Variant
      V = v1
      v1 = v2
      v2 = V
   End Sub
   
' Swaps 2 longs.
Private Sub SwapLongs(ByRef i1 As Long, _
                      ByRef i2 As Long)
   Dim i As Long
   i = i1
   i1 = i2
   i2 = i
End Sub
   
' Swaps 2 strings.
Private Sub SwapStrings(ByRef s1 As String, _
                        ByRef s2 As String)
   Dim i As Long
   ' StrPtr() returns 0 (NULL) if string is not initialized
   ' But StrPtr() is 5% faster than using CopyMemory, so I used that workaround, which is safe and fast.
   i = StrPtr(s1)

   If i = 0 Then CopyMemory ByVal VarPtr(i), ByVal VarPtr(s1), 4
   CopyMemory ByVal VarPtr(s1), ByVal VarPtr(s2), 4
   CopyMemory ByVal VarPtr(s2), i, 4
End Sub
   
' Fast hash algorithm.
Private Function GetFastXorHash(ByVal sString As String, _
                                Optional ByVal iLenToHash As Long = -1) As Long
   Dim i         As Long
   Dim iUBound   As Long
   Dim iBuffer() As Long

   If sString = vbNullString Then GetFastXorHash = -1: Exit Function
   If iLenToHash = -1 Then iLenToHash = Len(sString)
   If iLenToHash > Len(sString) Then iLenToHash = Len(sString)
   iUBound = iLenToHash \ 4 + 1 ' +1 to be sure
   ReDim iBuffer(iUBound)
   CopyMemory iBuffer(0), ByVal sString, iLenToHash

   For i = 0 To iUBound
      GetFastXorHash = GetFastXorHash Xor iBuffer(i) Xor i
   Next i

   GetFastXorHash = GetFastXorHash And &H7FFFFFFF
End Function
#End If '// mdlMarton_Loadable
'Public Sub Benchmark(Optional ByVal n As Long = 10000, Optional ByVal iLenString As Long = 100)
'   Dim iArray() As Long
'   Dim sArray() As String
'   Dim iIndex() As Long
'   Dim i As Long
'   Dim j As Long
'   Dim s As String
'   Dim iTime As Single
'
'   Randomize Timer
'
'   ReDim iArray(n)
'   ReDim sArray(n)
'   ReDim iIndex(0)
'
'   s = s & "Executing n(" & n & ") iterations. Len() String = " & iLenString & vbCrLf
'   's = s & "BEFORE:" & vbTab & "iArray()" & vbTab & "sArray()" & vbCrLf & vbCrLf
'   For i = 0 To n
'      iArray(i) = CLng(Rnd * n)
'
'      For j = 1 To iLenString
'         sArray(i) = sArray(i) & Chr(CLng(Rnd * 32 + 64))
'      Next j
'
'     's = s & vbTab & iArray(i) & vbTab & sArray(i) & vbCrLf
'   Next i
'
'   iTime = Timer
'
'   ' PLACE FUNCTION BELOW
'   ' ////////////////////////
''   TriQuickSortLong iArray
''   TriQuickSortString sArray
''   ShellSortLong iArray
''   ShellSortString sArray
'   ' ////////////////////////
'
'   iTime = Timer - iTime
'
'   's = s & vbCrLf & "AFTER:" & vbTab & "iArray()" & vbTab & "sArray()" & vbCrLf & vbCrLf
'   'For i = 0 To n
'   '   s = s & vbTab & iArray(i) & vbTab & sArray(i) & vbCrLf
'   'Next i
'
'   's = s & vbCrLf & "SORT CHECK: iArray() -> " & isSortedLongArray(iArray) & " // sArray() -> " & isSortedStringArray(sArray) & vbCrLf
'   s = s & "Execution time: " & Format(iTime, "#0.0000") & " secs."
'
'   MsgBox s
'End Sub
