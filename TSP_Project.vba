
Sub TSP()
    Dim Size As Integer
    Dim distance() As Single Dim traveled() As Boolean
    Dim Route() As Integer
    Dim totalDistance As Single
    Dim MatrixYcounter As Integer
    Dim StartPoint As Integer
    Dim EndPoint As Integer
    Dim minDistance As Single 'just to use in the distance loop 12. 
    Dim i As Integer, j As Integer, k As Integer, l As Integer

    Dim x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer 14. 
    Dim temp_dist As Single
    Dim coords As Range 'this is the table of coordinates
    
    ' 2 opt variables

    Dim Newroute() As Integer
    Dim Temproute1 As Integer
    Dim Temproute2 As Integer
    Dim Newroutedistance As Single
    Dim Iterationcounter As Integer
    Dim Iterationflag As Boolean

    ' determining how many cities are there?
    Size = Range(Range("a1").Offset(1, 0), Range("a1").Offset(1, 0).End(xlDown)).Rows.Count
    'now that we know the number of cities, redimension distance array
    ReDim distance(1 To Size, 1 To Size)
    'take the coordinates as a range
    Set coords = Range(Range("a2"), Range("a2").End(xlDown)).Resize(, 3)

    'control not any two nodes are the same node
    For i=1 To Size
        For j = i + 1 To Size
            If Range("b1").Offset(i).Value = Range("b1").Offset(j).Value Then
                If Range("c1").Offset(i).Value = Range("c1").Offset(j).Value Then
                    MsgBox "You have entered a node more than once"
                    Exit Sub 
                End If
            End If 
        Next j
    Next i


    'put in the first arm of the matrix
    Range("H3") = "City" 
    Range("H3").Font.Bold = True
    Range("H1") = "Distance Matrix"
    Range("H1").Font.Bold = True 
    With Range("H3")
    For i = 1 To Size
        .Offset(i, 0) = i
        .Offset(i, 0).Font.Bold = True
    Next
    'second arm of the matrix
    For j=1 To Size
        .Offset(0, j) = j
        .Offset(0, j).Font.Bold = True
    Next

    'fill it in with distances
    For i=1 To Size 
        For j = 1 To Size
        'the default value is 0
            If i = j Then
                Range("H3").Offset(i, j) = 0
                'otherwise look for euclidean distance
            Else
                'search for the coordinates for each value
                x1 = WorksheetFunction.VLookup(i, coords, 2, False) 'x of i
                y1 = WorksheetFunction.VLookup(i, coords, 3, False) 'y of i
                x2 = WorksheetFunction.VLookup(j, coords, 2, False) 'x of j
                y2 = WorksheetFunction.VLookup(j, coords, 3, False) 'y of j
                temp_dist = Sqr(((x1 - x2) ^ 2) + ((y1 - y2) ^ 2))
                'reading the distance
                distance(i, j) = temp_dist
                Range("H3").Offset(i, j) = temp_dist
            End If
        Next
    Next
    End With

    'Array where route will be stored. Starts and ends in City 1
    ReDim Route(1 To Size + 1)
    Route(1) = 1
    Route(Size + 1) = Route(1)

    'Boolean array indicating whether each city was already visited or not. Initialize all cities (except City 1) to False 
    ReDim traveled(1 To Size) 
    traveled(1) = True 
    For i=2 To Size
        traveled(i) = False 
    Next
    
    'Total distance traveled is initially 0. Initial current city is City 1
    totalDistance = 0
    StartPoint = 1
    For MatrixYcounter = 2 To Size 
        'initialize maxDistance to 0

        minDistance = 9999999 
        For i = 1 To Size
            If i <> StartPoint And Not traveled(i) Then
                If distance(StartPoint, i) < minDistance Then
                EndPoint = i
                minDistance = Range("H3").Offset(StartPoint, i)
                End If
            End If
        Next i
        'store the next city to be visited in the route array
        Route(MatrixYcounter) = EndPoint
        traveled(EndPoint) = True
        'update total distance travelled


        totalDistance = totalDistance + minDistance
        'update current city
        StartPoint = EndPoint 
    Next MatrixYcounter

    'Update total distance traveled with the distance between the last city visited and the initial city, City 1.
    totalDistance = totalDistance + distance(StartPoint, 1)

    'Print Results
    With Range("A2").Offset(Size + 5, 0)
        .Offset(0, 0).Value = "Nearest neighbor route"
        .Offset(1, 0).Value = "Stop #"
        .Offset(1, 1).Value = "City"
        
        For MatrixYcounter = 1 To Size + 1
            .Offset(MatrixYcounter + 1, 0).Value = MatrixYcounter
            .Offset(MatrixYcounter + 1, 1).Value = Route(MatrixYcounter)
        Next MatrixYcounter
        
        .Offset(Size + 4, 0).Value = "Total distance is " & totalDistance 

        For i=1 To Size+1
            .Offset(Size + 5 + i, 0).Value = Application.WorksheetFunction.VLookup(Route(i), coords, 2, True)
            .Offset(Size + 5 + i, 1).Value = Application.WorksheetFunction.VLookup(Route(i), coords, 3, True)
        Next i

        .Offset(Size + 6).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        ActiveSheet.Shapes.AddChart2(240, xlXYScatterLines).Select
        
        ActiveChart.ChartTitle.Text = "Initial basic feasible route by Nearest Neighbour" 153.
    End With

    ReDim Newroute(1 To Size + 1)
    Iterationcounter = 0 4.
    Iterationflag = False 6.

    Fori=1ToSize
        Newroute(i) = Route(i)
    Next i

    repatfromthestart:
    Iterationflag = False

    Fori=1ToSize
        For j = i + 1 To Size
            Newroute(i) = Route(j)
            Newroute(j) = Route(i)
            Newroute(Size + 1) = Newroute(1)
            For k = 1 To (j - i)
                Newroute(i + k) = Route(j - k)
            Next k

            For k = 1 To Size
                Newroutedistance = Newroutedistance + distance(Newroute(k), Newroute(k + 1))
            Next k 


            If Newroutedistance < totalDistance Then 31.

                Fork=1ToSize+1
                    Route(k) = Newroute(k)
                Next k


                totalDistance = Newroutedistance
                Iterationcounter = Iterationcounter + 1
                Iterationflag = True

                With Range("A2").Offset(Size + 5, 3 * Iterationcounter) 
                    .Offset(0, 0).Value = "Nearest neighbor route" 
                    .Offset(1, 0).Value = "Stop #"
                    .Offset(1, 1).Value = "City"
                    For MatrixYcounter = 1 To Size + 1
                        .Offset(MatrixYcounter + 1, 0).Value = MatrixYcounter .Offset(MatrixYcounter + 1, 1).Value = Route(MatrixYcounter)
                    Next MatrixYcounter
                    
                    .Offset(Size + 4, 0).Value = "Total distance is " & totalDistance


                    .Offset(Size + 5, 0).Value = "Iterationcounter = " & Iterationcounter 
                    For l = 1 To Size + 1
                        .Offset(Size + 5 + l, 0).Value = Application.WorksheetFunction.VLookup(Route(l), coords, 2, True) 
                        .Offset(Size + 5 + l, 1).Value = Application.WorksheetFunction.VLookup(Route(l), coords, 3, True)
                    Next l

                    .Offset(Size + 6).Select

                    Range(Selection, Selection.End(xlToRight)).Select 
                    Range(Selection, Selection.End(xlDown)).Select 
                    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLines).Select
                    ActiveChart.ChartTitle.Text = "Iteration" & Iterationcounter
                End With 
            End If
            Newroutedistance = 0
            If Iterationflag Then GoTo repatfromthestart 
        Next j
        
        For k = 1 To Size + 1 
            Newroute(k) = Route(k)
        Next k 
    Next i