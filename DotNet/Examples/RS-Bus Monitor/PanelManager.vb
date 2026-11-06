Option Explicit On
' 2026-11: VB.NET version of the class PanelManager of the VB6 program RS-Bus Monitor (PanelManager.cls).
' The layout, the margins and the colours are the same as in VB6. Differences with VB6:
' - VB6 draws directly in the PictureBox (AutoRedraw). Here the PictureBox is drawn in its Paint event (DrawAll);
'   a change of a LED or of a square asks for a new drawing (Invalidate).
' - The text is measured and drawn with TextRenderer (VB6: TextWidth, TextHeight and Print).
' - The enums are outside the class, and are used with their name (for example BlockStateType.WISSEL).

' ======= Enums voor betere leesbaarheid =======
Public Enum LedStateType
    LED_OFF = 0
    LED_ON = 1
    LED_NOT_CONNECTED = 2
End Enum

Public Enum BlockStateType
    NOT_CONNECTED = 0
    CENTRALE = 1
    WISSEL = 2
    BEZETMELDING = 3
End Enum

Public Class PanelManager

    ' ======= Configuratie voor de inhoud van picPanel =======
    Private Const NUM_BLOCKS As Integer = 4
    Private Const ROWS_PER_BLOCK As Integer = 32
    Private Const BITS_PER_ADDR As Integer = 8
    ' Marges in pixels
    Private Const MARGIN_L As Integer = 12
    Private Const MARGIN_T As Integer = 12
    Private Const MARGIN_R As Integer = 12
    Private Const MARGIN_B As Integer = 12
    ' Extra offsets (specifiek voor de inhoud van picPanel)
    Private Const LABEL_OFFSET As Integer = 8
    Private Const MIDGROUP_GAP As Integer = 6

    ' ======= Stijl / Kleuren =======
    Private COLOR_BACKGROUND As Color
    Private COLOR_LABEL As Color
    Private COLOR_LED_ON As Color
    Private COLOR_LED_OFF As Color
    Private COLOR_LED_NC As Color
    Private COLOR_LED_EDGE As Color
    Private COLOR_LED_GLOW As Color
    Private COLOR_BLOCK_NC As Color
    Private COLOR_CENTALE_NC As Color
    Private COLOR_WISSEL_NC As Color
    Private COLOR_BEZETMELDING As Color

    ' ======= Layout-variabelen =======
    Private m_RowH As Single
    Private m_D As Single
    Private m_R As Single
    Private m_Spacing As Single
    Private m_BlockGap As Single
    Private m_LabelW As Single
    Private m_BlockW As Single
    Private m_BlockPad As Single
    Private m_Sq As Single

    ' ======= State =======
    Private LedState(127, 7) As LedStateType
    Private blockState(127) As BlockStateType

    ' ======= Centers voor LEDs en blokjes =======
    Private m_Cx(,) As Single
    Private m_Cy(,) As Single
    Private m_Bx() As Single
    Private m_By() As Single

    ' ======= Referentie naar picPanel =======
    Private WithEvents m_PicPanel As PictureBox
    Private m_Font As Font                  ' VB6: the font of picPanel

    ' ======= Public Methods =======

    ' Initialiseert de PanelManager met een PictureBox.
    ' @param picPanel: De PictureBox waarin getekend wordt.
    Public Sub Init(picPanel As PictureBox)
        m_PicPanel = picPanel
        InitializeColors()
        InitializeFont()
        InitializeState()
        RebuildLayoutAndRedraw()
    End Sub


    ' Event welke de 4 LEDs toont die bij een bepaalde feedback nibble behoren
    ' @param address: 1..128
    ' @param nibble: 0 / 1 (0 = data is for bits 0..3 / 1 = data is for bits 4..7)
    ' @param data: 0..15
    Public Sub SetFeedbackNibble(ByVal address As Integer, ByVal nibble As Byte, ByVal data As Byte, ByVal bState As BlockStateType)
        If address < 1 Or address > 128 Then Exit Sub
        If nibble < 0 Or nibble > 1 Then Exit Sub
        If data < 0 Or data > 15 Then Exit Sub
        address = address - 1                   ' Internally addresses are from 0..127
        ' Ik moet alle vier de bits afzonderlijk testen en een boolean zetten als er iets wordt aangepast
        ' Als Nibble = 0, dan heb ik bits 0..3. Als Nibble = 1, dan heb ik bits 4..7
        Dim refreshNeeded As Boolean
        refreshNeeded = False
        Dim i As Byte                           ' Counter in the For statement
        Dim value As LedStateType               ' becomes 0 (LED_OFF) or 1 (LED_ON)
        Dim offset As Byte                      ' becomes 0 or 4
        If nibble = 0 Then offset = 0 Else offset = 4
        For i = 0 To 3
            value = CType(bitValue(data, i), LedStateType)
            If LedState(address, i + offset) <> value Then
                LedState(address, i + offset) = value
                refreshNeeded = True
            End If
        Next
        ' Kleur nu het blokje voor de LEDs
        If blockState(address) <> bState Then
            blockState(address) = bState
            refreshNeeded = True
        End If
        If refreshNeeded Then m_PicPanel.Invalidate()
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    ' bit manipulation function: returns 0 or 1
    Private Function bitValue(ByVal data As Byte, ByVal index As Byte) As Byte
        Return CByte((data >> index) And 1)
    End Function

    ' Stelt de staat van een LED in.
    ' @param addr: Adres (1..128).
    ' @param bit: Bit (0..7).
    ' @param state: Nieuwe staat (LED_OFF, LED_ON, LED_NOT_CONNECTED).
    Public Sub SetLed(ByVal addr As Integer, ByVal bit As Integer, ByVal state As LedStateType)
        addr = addr - 1
        If addr < 0 Or addr > 127 Then Exit Sub
        If bit < 0 Or bit > 7 Then Exit Sub
        If LedState(addr, bit) <> state Then
            LedState(addr, bit) = state
            m_PicPanel.Invalidate()
        End If
    End Sub

    ' Stelt de staat van een Blokje in.
    ' @param addr: Adres (1..128).
    ' @param state: Nieuwe staat (NOT_CONNECTED, CENTRALE, WISSEL, BEZETMELDING).
    Public Sub SetBlock(ByVal addr As Integer, ByVal state As BlockStateType)
        addr = addr - 1
        If addr < 0 Or addr > 127 Then Exit Sub
        If blockState(addr) <> state Then
            blockState(addr) = state
            m_PicPanel.Invalidate()
        End If
    End Sub

    ' Hertekent de layout en inhoud van picPanel.
    Public Sub Resize()
        RebuildLayoutAndRedraw()
    End Sub

    ' ======= Private Methods =======

    ' Initialiseert de kleuren.
    Private Sub InitializeColors()
        COLOR_BACKGROUND = Color.FromArgb(240, 240, 240)
        COLOR_LABEL = Color.FromArgb(0, 0, 255)
        COLOR_LED_ON = Color.FromArgb(220, 30, 30)
        COLOR_LED_OFF = Color.FromArgb(192, 242, 154)
        COLOR_LED_NC = Color.FromArgb(220, 220, 220)
        COLOR_LED_EDGE = Color.FromArgb(100, 100, 100)
        COLOR_LED_GLOW = Color.FromArgb(255, 200, 200)
        ' TODO: Betere kleuren geven
        COLOR_BLOCK_NC = Color.FromArgb(220, 220, 220)
        COLOR_CENTALE_NC = Color.FromArgb(255, 0, 0)
        COLOR_WISSEL_NC = Color.FromArgb(0, 255, 0)
        COLOR_BEZETMELDING = Color.FromArgb(0, 0, 255)
    End Sub

    ' Initialiseert het lettertype voor labels.
    Private Sub InitializeFont()
        m_Font = New Font("Microsoft Sans Serif", 10, FontStyle.Bold)
    End Sub

    ' Initialiseert de beginstate.
    Private Sub InitializeState()
        Dim a As Integer, k As Integer
        For a = 0 To 127
            blockState(a) = BlockStateType.NOT_CONNECTED
            For k = 0 To 7
                LedState(a, k) = LedStateType.LED_NOT_CONNECTED
            Next k
        Next a
    End Sub

    ' Herberekent de layout en tekent alles opnieuw.
    Private Sub RebuildLayoutAndRedraw()
        ComputeLayout(m_PicPanel.ClientSize.Width, m_PicPanel.ClientSize.Height)
        PrecomputeCenters()
        m_PicPanel.Invalidate()
    End Sub

    ' VB6: TextWidth and TextHeight of picPanel
    Private Function TextSize(ByVal text As String) As Size
        Return TextRenderer.MeasureText(text, m_Font, Size.Empty, TextFormatFlags.NoPadding)
    End Function

    ' Bereken de layout van het paneel.
    ' @param W: Breedte van picPanel.
    ' @param H: Hoogte van picPanel.
    Private Sub ComputeLayout(ByVal W As Single, ByVal H As Single)
        Dim contentW As Single, contentH As Single
        contentW = W - (MARGIN_L + MARGIN_R)
        contentH = H - (MARGIN_T + MARGIN_B)
        If contentW < 100 Or contentH < 100 Then Exit Sub

        ' Bereken de breedte voor de labelkolom
        m_LabelW = TextSize("128").Width + LABEL_OFFSET + 6

        ' Bereken de hoogte per rij
        m_RowH = CalculateRowHeight(contentH)

        ' Bereken de diameter van de bits en de spacing
        m_D = CalculateBitDiameter(m_RowH)
        m_R = m_D / 2
        m_Spacing = CalculateSpacing(m_D)

        ' Bereken de grootte van de blokjes
        m_Sq = CalculateBlockSize(m_RowH, m_R)

        ' Bereken de breedte van een blok
        CalculateBlockWidth(contentW)
    End Sub

    ' Bereken de hoogte per rij.
    Private Function CalculateRowHeight(ByVal contentH As Single) As Single
        CalculateRowHeight = Int(contentH / ROWS_PER_BLOCK)
        If CalculateRowHeight < 8 Then CalculateRowHeight = 8
    End Function

    ' Bereken de diameter van de bits.
    Private Function CalculateBitDiameter(ByVal rowH As Single) As Single
        CalculateBitDiameter = rowH - 4
        If CalculateBitDiameter < 6 Then CalculateBitDiameter = 6
    End Function

    ' Bereken de spacing tussen bits.
    Private Function CalculateSpacing(ByVal bitDiameter As Single) As Single
        CalculateSpacing = Math.Round(bitDiameter / 3)
        If CalculateSpacing < 2 Then CalculateSpacing = 2
    End Function

    ' Bereken de grootte van de blokjes.
    Private Function CalculateBlockSize(ByVal rowH As Single, ByVal radius As Single) As Single
        CalculateBlockSize = rowH / 2
        If CalculateBlockSize < 2 Then CalculateBlockSize = 2
        If CalculateBlockSize > (radius / 3) Then CalculateBlockSize = radius / 3
    End Function

    ' Bereken de breedte van een blok.
    Private Sub CalculateBlockWidth(ByVal contentW As Single)
        m_BlockGap = 24
        m_BlockPad = m_Spacing / 4
        If m_BlockPad < 1 Then m_BlockPad = 1

        ' Iteratief de breedte van een blok berekenen
        Do
            m_BlockW = m_LabelW + (BITS_PER_ADDR * m_D) +
                       ((BITS_PER_ADDR - 1) * m_Spacing) + MIDGROUP_GAP + m_BlockPad + m_Sq
            Dim totalW As Single
            totalW = (NUM_BLOCKS * m_BlockW) + ((NUM_BLOCKS - 1) * m_BlockGap)
            If totalW <= contentW Then Exit Do
            ' Pas de afmetingen aan als het niet past
            If m_D > 6 Then m_D = m_D - 1
            m_R = m_D / 2
            m_Spacing = CalculateSpacing(m_D)
            If m_BlockGap > 12 Then m_BlockGap = m_BlockGap - 1
            If m_D <= 6 And m_BlockGap <= 12 Then Exit Do
        Loop
    End Sub

    ' Bepaal de exacte pixelcoördinaten voor LEDs en blokjes.
    Private Sub PrecomputeCenters()
        Dim addr As Integer, b As Integer, r As Integer, k As Integer
        Dim baseX As Single, baseY As Single
        Dim ledOffsetX As Single
        ledOffsetX = m_BlockPad + m_Sq * 3

        ReDim m_Cx(NUM_BLOCKS * ROWS_PER_BLOCK - 1, BITS_PER_ADDR - 1)
        ReDim m_Cy(NUM_BLOCKS * ROWS_PER_BLOCK - 1, BITS_PER_ADDR - 1)
        ReDim m_Bx(127)
        ReDim m_By(127)

        For b = 0 To NUM_BLOCKS - 1
            baseX = MARGIN_L + b * (m_BlockW + m_BlockGap)
            For r = 0 To ROWS_PER_BLOCK - 1
                baseY = MARGIN_T + r * m_RowH
                addr = b * ROWS_PER_BLOCK + r

                ' Blokjes voor terugmeld type
                m_Bx(addr) = baseX + m_LabelW + m_BlockPad + m_Sq / 2
                m_By(addr) = baseY + m_RowH / 2

                ' LEDs
                For k = 0 To BITS_PER_ADDR - 1
                    Dim x As Single, y As Single
                    x = baseX + m_LabelW + ledOffsetX + k * (m_D + m_Spacing) + m_R
                    If k >= 4 Then x = x + MIDGROUP_GAP
                    y = baseY + (m_RowH / 2)
                    m_Cx(addr, k) = x
                    m_Cy(addr, k) = y
                Next k
            Next r
        Next b
    End Sub

    ' VB6: the drawing is made in DrawAll, DrawBlock and DrawLamp. Here the whole panel is drawn after each change
    Private Sub m_PicPanel_Paint(sender As Object, e As PaintEventArgs) Handles m_PicPanel.Paint
        DrawAll(e.Graphics)
    End Sub

    ' Tekent alles in picPanel.
    Private Sub DrawAll(ByVal g As Graphics)
        ' Wis de achtergrond
        g.Clear(COLOR_BACKGROUND)
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

        ' Teken alle adressen, labels, blokjes en LEDs
        Dim addr As Integer, r As Integer, k As Integer
        Dim labelText As String, labelX As Single, labelY As Single
        Dim b As Integer

        For addr = 0 To (NUM_BLOCKS * ROWS_PER_BLOCK) - 1
            b = addr \ ROWS_PER_BLOCK
            r = addr Mod ROWS_PER_BLOCK

            ' Teken labels
            If (r Mod 4) = 0 Then
                labelX = (MARGIN_L + b * (m_BlockW + m_BlockGap)) + LABEL_OFFSET
                labelY = MARGIN_T + r * m_RowH + (m_RowH \ 2) - (TextSize("0").Height \ 2)
                labelText = CStr(addr + 1)
                TextRenderer.DrawText(g, labelText, m_Font, New Point(CInt(labelX), CInt(labelY)), COLOR_LABEL, TextFormatFlags.NoPadding)
            End If

            ' Teken blokjes
            DrawBlock(g, addr, blockState(addr))

            ' Teken LEDs
            For k = 0 To BITS_PER_ADDR - 1
                DrawLamp(g, m_Cx(addr, k), m_Cy(addr, k), LedState(addr, k))
            Next k
        Next addr
    End Sub

    ' Tekent een blokje voor een adres.
    ' @param addr: Het adres (0..127).
    Private Sub DrawBlock(ByVal g As Graphics, ByVal addr As Integer, state As BlockStateType)
        If addr < 0 Or addr > 127 Then Exit Sub
        Dim fillC As Color
        Select Case state
            Case BlockStateType.NOT_CONNECTED : fillC = COLOR_BLOCK_NC
            Case BlockStateType.CENTRALE : fillC = COLOR_CENTALE_NC
            Case BlockStateType.WISSEL : fillC = COLOR_WISSEL_NC
            Case Else : fillC = COLOR_BEZETMELDING
        End Select
        Dim bx As Single, by As Single, s As Single
        bx = m_Bx(addr)
        by = m_By(addr)
        s = m_Sq
        Using brush As New SolidBrush(fillC)
            g.FillRectangle(brush, bx - s / 2, by - s / 2, s, s)
        End Using
    End Sub

    ' Tekent een LED-lampje.
    ' @param cx: X-coördinaat van het midden.
    ' @param cy: Y-coördinaat van het midden.
    ' @param state: Staat van de LED (LED_OFF, LED_ON, LED_NOT_CONNECTED).
    Private Sub DrawLamp(ByVal g As Graphics, ByVal cx As Single, ByVal cy As Single, ByVal state As LedStateType)
        Dim fillC As Color
        Select Case state
            Case LedStateType.LED_ON : fillC = COLOR_LED_ON
            Case LedStateType.LED_NOT_CONNECTED : fillC = COLOR_LED_NC
            Case Else : fillC = COLOR_LED_OFF
        End Select

        Circle(g, cx, cy, m_R, fillC, COLOR_LED_EDGE)

        ' Teken de glans als de LED aan staat en een subtiele glans toe als de LED uit is
        If state = LedStateType.LED_ON And m_R >= 6 Then
            Circle(g, cx - m_R / 3, cy - m_R / 3, m_R / 3, COLOR_LED_GLOW, fillC)
        ElseIf state = LedStateType.LED_OFF And m_R >= 6 Then
            Circle(g, cx - m_R / 3, cy - m_R / 3, m_R / 4, Color.FromArgb(200, 200, 200), fillC) ' Kleinere glans
        End If
    End Sub

    ' VB6: FillColor, FillStyle = vbFSSolid and Circle (cx, cy), radius, edgeColor
    Private Sub Circle(ByVal g As Graphics, ByVal cx As Single, ByVal cy As Single, ByVal radius As Single, ByVal fillC As Color, ByVal edgeC As Color)
        Using brush As New SolidBrush(fillC)
            g.FillEllipse(brush, cx - radius, cy - radius, 2 * radius, 2 * radius)
        End Using
        Using pen As New Pen(edgeC)
            g.DrawEllipse(pen, cx - radius, cy - radius, 2 * radius, 2 * radius)
        End Using
    End Sub
End Class
