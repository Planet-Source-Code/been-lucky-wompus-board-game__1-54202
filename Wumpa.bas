Attribute VB_Name = "Module1"
Global Game(30) As String * 144, GameDeck As String * 30, GameNo As Integer ' game & deck variables
Global Tile(12, 12) As Integer, Floor, Wall, Stone, Ball, MaxX, MaxY        ' game board variables
Global Moved As Boolean, Direction As String                                ' action variables
Global BlockerMode As Integer, Colour(8) As Long                            ' custom setting variables


Sub Main()
    
    Randomize Timer
    LoadGames           ' Load the 30 games and create a deck of games(mixed)
    SetColourScheme 0   ' Default scheme
    Initialise          '
    Form2.Show 1        ' A modal form ... unloads when user presses enter key
                        '           or does a mouse click on Form2
    
    
    Form1.Show          ' Lets start a new game
    StartNewGame
    Form1.Picture1.Cls
    Showblocks
                        
                        
                        
    Do                  ' The game loop. We play the game till it ends
    
        If Direction <> "" Then MoveBlocks
        DoEvents
        Loop
        
        
        
        

End Sub

Sub Initialise()
    
    Floor = 0
    Wall = 1
    Stone = 2
    Ball = 3
    MaxX = 12
    MaxY = 12
    Direction = ""
    Load_A_Game             ' will load the next game in a sequence

    With Form1.Picture1
        .ScaleWidth = 12
        .ScaleHeight = 12
        .Cls
        End With
    
End Sub
Sub StartNewGame()
    
    Initialise

End Sub
Sub MoveBlocks()

    Moved = False
    
    Select Case Direction
        Case "N": GoSub MoveUp
        Case "S": GoSub MoveDown
        Case "W": GoSub MoveLeft
        Case "E": GoSub MoveRight
        End Select
    
        Exit Sub        ' moving block over so we get out


'----  Internal Subroutines

MoveLeft:
        Moved = False
        For X = 1 To MaxX - 1: For Y = 1 To MaxY
            
            If Tile(X, Y) = Floor And Tile(X + 1, Y) = Ball Then
            Tile(X, Y) = Ball: Tile(X + 1, Y) = Floor: Moved = True
            End If
            
            
            If Tile(X, Y) = Floor And Tile(X + 1, Y) = Stone Then
            Tile(X, Y) = Stone: Tile(X + 1, Y) = Floor: Moved = True
            If GameSolved Then Exit Sub
            End If
            
        Next: Next
        
        
        If Moved = False Then
                Return
        Else:   Showblocks              ' if moved show updated position of blocks
                GoTo MoveLeft           ' move blocks till all blocks moved
                End If
        
MoveRight:
        Moved = False
        For X = MaxX To 2 Step -1: For Y = 1 To MaxY
            
            If Tile(X, Y) = Floor And Tile(X - 1, Y) = Ball Then
            Tile(X, Y) = Ball: Tile(X - 1, Y) = Floor: Moved = True
            End If
            
            
            If Tile(X, Y) = Floor And Tile(X - 1, Y) = Stone Then
            Tile(X, Y) = Stone: Tile(X - 1, Y) = Floor: Moved = True
            If GameSolved Then Exit Sub
            End If
        
        
        Next: Next
        
        
        If Moved = False Then
                Return
        Else:   Showblocks
                GoTo MoveRight
                End If
        
MoveDown:
        Moved = False
        For Y = 1 To MaxY - 1: For X = 1 To MaxX
        
            If Tile(X, Y) = Ball And Tile(X, Y + 1) = Floor Then
            Tile(X, Y) = Floor: Tile(X, Y + 1) = Ball: Moved = True
            End If
            
            If Tile(X, Y) = Stone And Tile(X, Y + 1) = Floor Then
            Tile(X, Y) = Floor: Tile(X, Y + 1) = Stone: Moved = True
            If GameSolved Then Exit Sub
            End If
                    
        
        Next: Next
        
        
        
        If Moved = False Then
                Return
        Else:   Showblocks
                GoTo MoveDown
                End If
        
        

MoveUp:

        Moved = False
        For Y = MaxY To 2 Step -1: For X = 1 To MaxX
        
            
            If Tile(X, Y) = Ball And Tile(X, Y - 1) = Floor Then
            Tile(X, Y) = Floor: Tile(X, Y - 1) = Ball: Moved = True
            End If

            
            If Tile(X, Y) = Stone And Tile(X, Y - 1) = Floor Then
            Tile(X, Y) = Floor: Tile(X, Y - 1) = Stone: Moved = True
            If GameSolved Then Exit Sub
            End If
            
            
        Next: Next
        
        
        
        
        If Moved = False Then
                Return
        Else:   Showblocks
                GoTo MoveUp
                End If
            

        Direction = ""          '  ie reset ... we finished acting on direction


End Sub
Sub Showblocks()

    
    For X = 1 To MaxX: For Y = 1 To MaxY
    Select Case Tile(X, Y)
        Case Wall:
                    Form1.Picture1.ForeColor = Colour(1)
                    Select Case BlockerMode  ' 0 = box; 1 = lines
                        Case 0: Form1.Picture1.Line (X - 0.1, Y - 0.1)-(X - 0.9, Y - 0.9), Colour(2), BF
                        Case 1: Form1.Picture1.Line (X, Y)-(X - 1, Y - 1), Colour(2)
                                Form1.Picture1.Line (X, Y - 1)-(X - 1, Y), Colour(2)
                        
                        End Select
                        
        Case Floor: Form1.Picture1.FillColor = Colour(3)
                    Form1.Picture1.ForeColor = Colour(4)
                    Form1.Picture1.Circle (X - 0.5, Y - 0.5), 0.4
        Case Stone: Form1.Picture1.FillColor = Colour(5)
                    Form1.Picture1.ForeColor = Colour(6)
                    Form1.Picture1.Circle (X - 0.5, Y - 0.5), 0.4
        Case Ball:  Form1.Picture1.FillColor = Colour(8)
                    Form1.Picture1.ForeColor = Colour(7)
                    Form1.Picture1.Circle (X - 0.5, Y - 0.5), 0.4
        End Select
        
    Next: Next
    

    
End Sub
Sub SetColourScheme(N As Integer)

Red = RGB(255, 0, 0): Green = RGB(0, 255, 0): Blue = RGB(0, 0, 255)
Skyblue = RGB(0, 255, 255): Pink = RGB(255, 0, 255): Yellow = RGB(255, 255, 0)
White = RGB(255, 255, 255): Black = RGB(0, 0, 0):
Violet = RGB(127, 127, 255): Darkgreen = RGB(0, 209, 0): Darkblue = RGB(50, 90, 233)
Brown = RGB(224, 59, 44): Orangedot = RGB(255, 162, 83)
'PaleGreen =RGB(127, 255, 55): PaleYellowdot=RGB(255, 255, 87):YellowDots=RGB(240, 240, 160)


'   Use     Colours 1,2 for wall
'           Colours 3,4 for stone
'           Colour 5,6 for floor
'           Colour 7,8 for ball
    
    Select Case N
        
        Case 0:  ' This is the default color scheme
                
                Colour(1) = White
                Colour(2) = White
                Colour(3) = Form1.Picture1.BackColor
                Colour(4) = Form1.Picture1.BackColor
                Colour(5) = Black
                Colour(6) = Black
                Colour(7) = White
                Colour(8) = Green
 
        
        Case 1: ' Scheme A
        
                Colour(1) = Black
                Colour(2) = Red
                Colour(3) = Form1.Picture1.BackColor
                Colour(4) = Form1.Picture1.BackColor
                Colour(5) = White
                Colour(6) = Blue
                Colour(7) = Red
                Colour(8) = Green
        
        Case 2: ' Scheme B
                            
                Colour(1) = Black
                Colour(2) = Violet
                Colour(3) = Form1.Picture1.BackColor
                Colour(4) = Form1.Picture1.BackColor
                Colour(5) = White
                Colour(6) = Brown
                Colour(7) = Red
                Colour(8) = Green
               
               
        Case 3: ' Scheme C
                
                Colour(1) = Black
                Colour(2) = Darkgreen
                Colour(3) = Form1.Picture1.BackColor
                Colour(4) = Form1.Picture1.BackColor
                Colour(5) = Red
                Colour(6) = Blue
                Colour(7) = Blue
                Colour(8) = Green
        
        Case 4: ' Scheme D
        
                Colour(1) = Black
                Colour(2) = Orangedot
                Colour(3) = Form1.Picture1.BackColor
                Colour(4) = Form1.Picture1.BackColor
                Colour(5) = Pink
                Colour(6) = Blue
                Colour(7) = Blue
                Colour(8) = Green
        
        Case 5: ' Scheme E
        
                Colour(1) = Black
                Colour(2) = Violet
                Colour(3) = Form1.Picture1.BackColor
                Colour(4) = Form1.Picture1.BackColor
                Colour(5) = Orangedot
                Colour(6) = Blue
                Colour(7) = Blue
                Colour(8) = Green
    
    
    End Select


End Sub
Sub ShowAbout()

    N$ = Chr$(13) & Chr$(10)
    a$ = "For more information refer to file : README.TXT"
    'a$ = a$ & N$ & App.Path & "\Info.txt"
    a$ = a$ & N$ & N$ & "PRESS ENTER TO CONTINUE PLAY"

    Load Form2
    
    With Form2
        '.Height = .Height + .Label1.Height
        .Label1 = a$
        .Show 1
        End With
    
End Sub
Function GameSolved() As Boolean

    GameSolved = (Tile(12, 12) = Ball)

    If GameSolved Then
        Showblocks  ' update first
        Wants2Play = MsgBox("Congratulations !!  You solved the Wompus. Would you like to play again ?", 4)
        If (Wants2Play = 6) Then StartNewGame: Showblocks Else End
        End If
        
    
End Function
Sub LoadGames()

'   30 built in games

    Game(1) = "300212101020002222221021022000012212221101110221020202100012110110112102110022220111122100120222220200111000221101100120110110002020102122002110"
    Game(2) = "300022121200000001120001212201010020200022011101000120220022200202202220222222002021110110002020122212101112010011210110100022110220022221111100"
    Game(3) = "300101011212002211221020010021211110010122122111222021210100201220110011211120221201221110201202121202102022220112221112102002020220121121012100"
    Game(4) = "321020110212210100120010020000200201202120220211020220010012002010000002001000100221020200100020222220220022222212112212101200112010210011010100"
    Game(5) = "301122000212201111020100210011002222020020100100002200000202110210110210210002121210022111121002020221212002121000201222011101221200221211111110"
    Game(6) = "311022211000002210112222020100122122120121022000102212220102110020212201221022220101210000001012100020220101110111100222120100111222001111122210"
    Game(7) = "322101101120020010100101012122201200222020112212212220211120211100012010111110120102112000222212022220012120212122200110012211200020102221122200"
    Game(8) = "300120221112212202012220210022201200202101120202202021120100000112001022211120122110100100202221122210122220101210210110002222212222211221011200"
    Game(9) = "312111220001220200221002112011021020022101020022201011021201211200120002012010011101120001121221021120200202001022212012210221001012211000120200"
    Game(10) = "321020122222200020212122212220222020202112002102002212220021011121121120202121221200122122100120222121002022202120202022002012022010202000011020"
    Game(11) = "322212202001101011222201121221120112011201212211110210001120100122022121002100110002220222200201122100110210102120200001200211102220000121101100"
    Game(12) = "300022221220202001021212010102001001112011100110210122002001120021112202121111222200101202012102220011221012010012111212000021221102110022211020"
    Game(13) = "300001200221011221202110000222022120212200211020111220000111211200020022011220220020012122221222022012020210100200022011020222101222220120221010"
    Game(14) = "320200120011121102010120120021101020211022021022202121200220002111200202200012112002111221000020011110122100001021002210220000201000220101001110"
    Game(15) = "321000012021020012210101220101121002020210122220002210011112010111020000120122021010121111011001120010021000012000111202212212102202022101220210"
    Game(16) = "320220010021002200100001220011112202211021120100221102000111121000111210221022021122102201222111201121210211210000121120011112122021011000002000"
    Game(17) = "302212001200100112201120100210010021011020222201211022201000120020202110011212020201000100021000111102212112202001202220202000201110111120202120"
    Game(18) = "310020201201211000222112020011001112212122210111110222002020101011210220112211210020212201001021121011101010222020220002220112102001211211202200"
    Game(19) = "321121100210120212101000201210212111220020122202121210011010112221202212002000122102101021210221201112202212212110222010011110102120222010002200"
    Game(20) = "300110102112221222222221222222200110000102000212000110201110112210102110011002120220010212100220210212212020112110110112100020020121021210202220"
    Game(21) = "310222002012002100021222212222021120211111210121100022020020011012210011210011221110120110011120212120022012021112021100220022010020222121202210"
    Game(22) = "302112101111121102020012001101121010000012120121210020020122112102210001120121121102110210000210220022000202122100202112001202122220211112010210"
    Game(23) = "321202020012022212011021210210122100010011010100102011022222212012110021012202010012100102121022021121010200220010200111022200110222222000010210"
    Game(24) = "311200000210202221000011220001120002002202121012120211002020200022010210001102102200020000020111010111120012100111120110101110220212220200202220"
    Game(25) = "322221000011212111001202000011020212122001200001121210001221210201222120210120002102010212102110020010100221112001022211001002000222000112112000"
    Game(26) = "322101110122102102121022020101012111021020200110120101212001210202121222120210122210201111200002112221101021222210120000100011022010022111102010"
    Game(27) = "301200010202102210021101202100220122201122022000100110000122012101102202010000000102200112020022120010112200000012012010102222202022220001211210"
    Game(28) = "312022022101201212111022021222020101002102101201002011011202221201022211120102100221111202020120112221012200212101000020101020021101201022101120"
    Game(29) = "322112221121202021011022000121212000100011122002112000022221212100110122010112020222001200122112010001112002020212221001201222021202110112221020"
    Game(30) = "302002200012002002012210120221110022001210102212012210210111100121111002110022001001220000210221010221012001000001221200000111120022120121111110"


    ' Creation of the random deck of games. Just a simple algorithm I thought up.
    
    For N = 1 To 30
        Select Case Int(Rnd * 5) + 1
        Case 1: a1$ = a1$ + Chr$(N)
        Case 2: a2$ = a2$ + Chr$(N)
        Case 3: a3$ = a3$ + Chr$(N)
        Case 4: a4$ = a4$ + Chr$(N)
        Case Else: a5$ = a5$ + Chr$(N)
        End Select
    Next
    
    GameDeck = a1$ & a2$ & a3$ & a4$ & a5$

End Sub

Sub Load_A_Game()
    
    ' Our deck has random mixed 30 games... from which we select next game
    If GameNo < 30 Then GameNo = GameNo + 1 Else GameNo = 1
    GurrentGame = Asc(Mid$(GameDeck, GameNo, 1))
    
    
    ' Now fill the grid tiles
    For X = 1 To MaxX: For Y = 1 To MaxY
        N = N + 1: p = Val(Mid$(Game(GurrentGame), N, 1))
        Tile(X, Y) = p
        Next: Next
    
End Sub

