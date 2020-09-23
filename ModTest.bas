Attribute VB_Name = "ModTest"
Attribute VB_HelpID = 2002
Sub main()

Dim PDFPrinter As New PDFPrinter
        
    PDFPrinter.PDFTitle = "Test"
    
    PDFPrinter.PDFFileName = App.Path & "\Defaut.pdf"

    PDFPrinter.PDFLoadAfm = App.Path & "\Fonts"
    PDFPrinter.PDFConfirm = False
    PDFPrinter.PDFView = True
    PDFPrinter.PDFFiligran = "P D F P r i n t e r   D e m o"
    
    PDFPrinter.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR
    PDFPrinter.PDFFormatPage = FORMAT_A4
    PDFPrinter.PDFOrientation = ORIENT_PORTRAIT
    PDFPrinter.PDFSetUnit = UNIT_PT
    PDFPrinter.PDFSetZoomMode = ZOOM_FULLWIDTH
    PDFPrinter.PDFSetLayoutMode = LAYOUT_DEFAULT
    PDFPrinter.PDFUseOutlines = True
    PDFPrinter.PDFUseThumbs = False
    
    PDFPrinter.PDFBeginDoc
        PDFPrinter.PDFSetLineStyle = pPDF_SOLID
        PDFPrinter.PDFSetLineWidth = 1
        PDFPrinter.PDFDrawLine 15, 20, 550, 20

        PDFPrinter.PDFSetBookmark "Signet 1", 0, 40
        PDFPrinter.PDFSetBookmark "Sous-Signet 2", 1, 60
        
        PDFPrinter.PDFSetFont FONT_ARIAL, 15, FONT_BOLD
        PDFPrinter.PDFSetTextColor = COLOR_NOIR
        PDFPrinter.PDFTextOut "Signet", 15, 40

        PDFPrinter.PDFSetFont FONT_ARIAL, 15, FONT_BOLD & FONT_UNDERLINE
        PDFPrinter.PDFSetTextColor = COLOR_ROUGE
        PDFPrinter.PDFTextOut "Sous-Signet 1", 15, 60

        PDFPrinter.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDFPrinter.PDFSetDrawColor = COLOR_ROUGE
        PDFPrinter.PDFSetTextColor = COLOR_BLANC
        PDFPrinter.PDFSetAlignement = ALIGN_FJUSTIFY
        PDFPrinter.PDFSetBorder = BORDER_ALL
        PDFPrinter.PDFSetFill = True
        PDFPrinter.PDFCell "Test Alignement Justifié Forcé", 15, 90, 400, 40, "http://yahoo.fr"

        PDFPrinter.PDFSetFont FONT_COURIER, 12, FONT_BOLD
        PDFPrinter.PDFSetDrawColor = COLOR_BLEU
        PDFPrinter.PDFSetTextColor = COLOR_BLANC
        PDFPrinter.PDFSetAlignement = ALIGN_CENTER
        PDFPrinter.PDFSetBorder = BORDER_NONE
        PDFPrinter.PDFSetFill = True
        PDFPrinter.PDFCell "Test Alignement au Centre", 15, 140, 400, 40

        PDFPrinter.PDFSetFont FONT_ARIAL, 12, FONT_BOLD & FONT_UNDERLINE
        PDFPrinter.PDFSetDrawColor = COLOR_JAUNE
        PDFPrinter.PDFSetTextColor = COLOR_NOIR
        PDFPrinter.PDFSetAlignement = ALIGN_LEFT
        PDFPrinter.PDFSetBorder = BORDER_ALL
        PDFPrinter.PDFSetFill = True
        PDFPrinter.PDFCell "Test Alignement à Gauche", 15, 190, 400, 40

        PDFPrinter.PDFSetLineColor = COLOR_VERT
        PDFPrinter.PDFSetLineWidth = 0.5
        PDFPrinter.PDFSetLineStyle = pPDF_DASHDOTDOT
        PDFPrinter.PDFDrawLineVer 15, 250, 100

        PDFPrinter.PDFSetLineColor = COLOR_ROUGE
        PDFPrinter.PDFSetLineWidth = 1
        PDFPrinter.PDFSetLineStyle = pPDF_DASH
        PDFPrinter.PDFDrawLineHor 30, 250, 100

        PDFPrinter.PDFSetTextColor = COLOR_NOIR
        PDFPrinter.PDFTextOut "Lien hypertext : ", 30, 280
        PDFPrinter.PDFLink 140, 280, "http://www.google.fr"

        PDFPrinter.PDFSetTextColor = COLOR_NOIR
        PDFPrinter.PDFSetRotation = 15
        PDFPrinter.PDFTextOut "Lien hypertext : ", 30, 330
        PDFPrinter.PDFLink 140, 330, "http://www.club-internet.fr"

        PDFPrinter.PDFSetDrawColor = COLOR_VERT
        PDFPrinter.PDFSetLineColor = COLOR_MAGENTA
        PDFPrinter.PDFSetLineStyle = pPDF_DASH
        PDFPrinter.PDFSetLineWidth = 0.5
        PDFPrinter.PDFSetDrawMode = DRAW_DRAWBORDER
        PDFPrinter.PDFDrawRectangle 15, 360, 400, 50, "Essai URL"

        PDFPrinter.PDFSetLineColor = COLOR_NOIR
        PDFPrinter.PDFSetLineStyle = pPDF_SOLID
        PDFPrinter.PDFSetLineWidth = 1
        PDFPrinter.PDFSetDrawMode = DRAW_NORMAL
        PDFPrinter.PDFDrawRectangle 15, 430, 400, 50

        PDFPrinter.PDFSetDrawColor = COLOR_JAUNE
        PDFPrinter.PDFSetLineColor = COLOR_CYAN
        PDFPrinter.PDFSetLineStyle = pPDF_DASH
        PDFPrinter.PDFSetLineWidth = 0.75
        PDFPrinter.PDFSetDrawMode = DRAW_DRAWBORDER
        PDFPrinter.PDFDrawPolygon Array(15, 520, 150, 520, 75, 570)

        PDFPrinter.PDFSetDrawColor = COLOR_MAGENTA
        PDFPrinter.PDFSetLineColor = COLOR_NONE
        PDFPrinter.PDFSetLineStyle = pPDF_DASHDOT
        PDFPrinter.PDFSetLineWidth = 1.25
        PDFPrinter.PDFSetDrawMode = DRAW_DRAW
        PDFPrinter.PDFDrawPolygon Array(200, 520, 300, 520, 250, 570)

        PDFPrinter.PDFSetDrawColor = COLOR_NONE
        PDFPrinter.PDFSetLineColor = COLOR_ROUGE
        PDFPrinter.PDFSetLineStyle = pPDF_DASHDOT
        PDFPrinter.PDFSetLineWidth = 0.5
        PDFPrinter.PDFSetDrawMode = DRAW_NORMAL
        PDFPrinter.PDFDrawPolygon Array(350, 520, 425, 520, 375, 570)

        PDFPrinter.PDFSetDrawColor = COLOR_CYAN
        PDFPrinter.PDFSetLineColor = COLOR_ROUGE
        PDFPrinter.PDFSetLineStyle = pPDF_DASH
        PDFPrinter.PDFSetLineWidth = 0.75
        PDFPrinter.PDFSetDrawMode = DRAW_DRAWBORDER
        PDFPrinter.PDFDrawPolygon Array(475, 520, 525, 520, 575, 570)

        PDFPrinter.PDFSetLineColor = COLOR_NONE
        PDFPrinter.PDFSetDrawColor = COLOR_VERT
        PDFPrinter.PDFSetLineStyle = pPDF_SOLID
        PDFPrinter.PDFSetLineWidth = 1
        PDFPrinter.PDFSetDrawMode = DRAW_DRAW
        PDFPrinter.PDFDrawEllipse 25, 645, 75, 25

        PDFPrinter.PDFSetLineColor = COLOR_BLEU
        PDFPrinter.PDFSetLineStyle = pPDF_DASH
        PDFPrinter.PDFSetLineWidth = 0.25
        PDFPrinter.PDFSetDrawMode = DRAW_NORMAL
        PDFPrinter.PDFDrawEllipse 110, 620, 75

        PDFPrinter.PDFSetDrawColor = COLOR_MAGENTA
        PDFPrinter.PDFSetLineColor = COLOR_CYAN
        PDFPrinter.PDFSetLineStyle = pPDF_DASH
        PDFPrinter.PDFSetLineWidth = 0.75
        PDFPrinter.PDFSetDrawMode = DRAW_DRAWBORDER
        PDFPrinter.PDFDrawEllipse 210, 620, 75

        PDFPrinter.PDFSetLineColor = COLOR_NONE
        PDFPrinter.PDFSetDrawColor = COLOR_ROUGE
        PDFPrinter.PDFSetLineStyle = pPDF_SOLID
        PDFPrinter.PDFSetLineWidth = 1
        PDFPrinter.PDFSetDrawMode = DRAW_DRAW
        PDFPrinter.PDFDrawEllipse 310, 620, 75

        PDFPrinter.PDFSetLineColor = COLOR_NOIR
        PDFPrinter.PDFSetLineStyle = pPDF_DASHDOTDOT
        PDFPrinter.PDFSetLineWidth = 1.75
        PDFPrinter.PDFSetDrawMode = DRAW_NORMAL
        PDFPrinter.PDFDrawEllipse 425, 620, 25, 75

        PDFPrinter.PDFSetDrawColor = COLOR_JAUNE
        PDFPrinter.PDFSetLineColor = COLOR_NOIR
        PDFPrinter.PDFSetLineStyle = pPDF_DASHDOT
        PDFPrinter.PDFSetLineWidth = 1.25
        PDFPrinter.PDFSetDrawMode = DRAW_DRAWBORDER
        PDFPrinter.PDFDrawEllipse 475, 645, 75, 25, "http://www.voila.fr"

        PDFPrinter.PDFSetFont FONT_ARIAL, 8, FONT_ITALIC
        PDFPrinter.PDFSetTextColor = COLOR_NOIR
        PDFPrinter.PDFTextOut "Page " & PDFPrinter.PDFPageNumber, _
                         PDFPrinter.PDFGetPageWidth / 2 - PDFPrinter.PDFGetStringWidth("Page " & PDFPrinter.PDFPageNumber), _
                         PDFPrinter.PDFGetPageHeight - PDFPrinter.PDFTextHeight

        PDFPrinter.PDFEndPage
        
        PDFPrinter.PDFNewPage

        PDFPrinter.PDFSetBookmark "Signet 2", 0, 40
        PDFPrinter.PDFSetBookmark "Sous-Signet 2", 1, 60
        
        PDFPrinter.PDFSetFont FONT_ARIAL, 15, FONT_BOLD
        PDFPrinter.PDFSetTextColor = COLOR_NOIR
        PDFPrinter.PDFTextOut "Signet 2", 15, 40

        PDFPrinter.PDFSetFont FONT_ARIAL, 15, FONT_BOLD & FONT_UNDERLINE
        PDFPrinter.PDFSetTextColor = COLOR_ROUGE
        PDFPrinter.PDFTextOut "Sous-Signet 2", 15, 60
        
        PDFPrinter.PDFSetDrawColor = COLOR_CYAN
        PDFPrinter.PDFSetLineColor = COLOR_MAGENTA
        PDFPrinter.PDFSetLineStyle = pPDF_SOLID
        PDFPrinter.PDFSetLineWidth = 0.65
        PDFPrinter.PDFSetDrawMode = DRAW_DRAW
        PDFPrinter.PDFDrawRectangle 15, 80, 400, 50

        PDFPrinter.PDFSetLineColor = COLOR_BLEU
        PDFPrinter.PDFSetLineStyle = pPDF_SOLID
        PDFPrinter.PDFSetLineWidth = 0.85
        PDFPrinter.PDFSetDrawMode = DRAW_NORMAL
        PDFPrinter.PDFDrawRectangle 15, 150, 400, 50

        PDFPrinter.PDFImage App.Path & "\NeO78.jpg", 15, 220, 150, 100, "http://www.google.fr"
        
        PDFPrinter.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
        PDFPrinter.PDFSetTextColor = COLOR_NOIR
        PDFPrinter.PDFSetDrawColor = COLOR_JAUNE
        PDFPrinter.PDFSetLineColor = COLOR_NOIR
        PDFPrinter.PDFSetAlignement = ALIGN_LEFT
        PDFPrinter.PDFSetBorder = BORDER_ALL
        PDFPrinter.PDFSetLineWidth = 1
        PDFPrinter.PDFSetFill = True
        PDFPrinter.PDFCell "Test Alignement à Gauche", 15, 350, 150, 20
        
        PDFPrinter.PDFSetFont FONT_ARIAL, 8, FONT_ITALIC
        PDFPrinter.PDFSetTextColor = COLOR_NOIR
        PDFPrinter.PDFTextOut "Page " & PDFPrinter.PDFPageNumber, _
                         PDFPrinter.PDFGetPageWidth / 2 - PDFPrinter.PDFGetStringWidth("Page " & PDFPrinter.PDFPageNumber), _
                         PDFPrinter.PDFGetPageHeight - PDFPrinter.PDFTextHeight
    
        PDFPrinter.PDFEndPage
        
        PDFPrinter.PDFNewPage

        PDFPrinter.PDFSetBookmark "Signet 2", 0, 40
        PDFPrinter.PDFSetBookmark "Sous-Signet 2", 1, 60
        
        PDFPrinter.PDFSetFont FONT_ARIAL, 15, FONT_BOLD
        PDFPrinter.PDFSetTextColor = COLOR_NOIR
        PDFPrinter.PDFTextOut "Signet 2", 15, 40

        PDFPrinter.PDFSetFont FONT_ARIAL, 15, FONT_BOLD & FONT_UNDERLINE
        PDFPrinter.PDFSetTextColor = COLOR_ROUGE
        PDFPrinter.PDFTextOut "Sous-Signet 2", 15, 60
        
        PDFPrinter.PDFSetDrawColor = COLOR_CYAN
        PDFPrinter.PDFSetLineColor = COLOR_MAGENTA
        PDFPrinter.PDFSetLineStyle = pPDF_SOLID
        PDFPrinter.PDFSetLineWidth = 0.65
        PDFPrinter.PDFSetDrawMode = DRAW_DRAW
        PDFPrinter.PDFDrawRectangle 15, 80, 400, 50

        PDFPrinter.PDFSetLineColor = COLOR_BLEU
        PDFPrinter.PDFSetLineStyle = pPDF_SOLID
        PDFPrinter.PDFSetLineWidth = 0.85
        PDFPrinter.PDFSetDrawMode = DRAW_NORMAL
        PDFPrinter.PDFDrawRectangle 15, 150, 400, 50

        PDFPrinter.PDFImage App.Path & "\NeO78.jpg", 15, 220, 150, 100, "http://www.google.fr"
        
        PDFPrinter.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
        PDFPrinter.PDFSetTextColor = COLOR_NOIR
        PDFPrinter.PDFSetDrawColor = COLOR_JAUNE
        PDFPrinter.PDFSetLineColor = COLOR_NOIR
        PDFPrinter.PDFSetAlignement = ALIGN_LEFT
        PDFPrinter.PDFSetBorder = BORDER_ALL
        PDFPrinter.PDFSetLineWidth = 1
        PDFPrinter.PDFSetFill = True
        PDFPrinter.PDFCell "Test Alignement à Gauche", 15, 350, 150, 20
        
        PDFPrinter.PDFSetFont FONT_ARIAL, 8, FONT_ITALIC
        PDFPrinter.PDFSetTextColor = COLOR_NOIR
        PDFPrinter.PDFTextOut "Page " & PDFPrinter.PDFPageNumber, _
                         PDFPrinter.PDFGetPageWidth / 2 - PDFPrinter.PDFGetStringWidth("Page " & PDFPrinter.PDFPageNumber), _
                         PDFPrinter.PDFGetPageHeight - PDFPrinter.PDFTextHeight
                         
        PDFPrinter.PDFEndPage
                                 
        PDFPrinter.PDFNewPage

        PDFPrinter.PDFSetBookmark "Signet 2", 0, 40
        PDFPrinter.PDFSetBookmark "Sous-Signet 2", 1, 60
        
        PDFPrinter.PDFSetFont FONT_ARIAL, 15, FONT_BOLD
        PDFPrinter.PDFSetTextColor = COLOR_NOIR
        PDFPrinter.PDFTextOut "Signet 2", 15, 40

        PDFPrinter.PDFSetFont FONT_ARIAL, 15, FONT_BOLD & FONT_UNDERLINE
        PDFPrinter.PDFSetTextColor = COLOR_ROUGE
        PDFPrinter.PDFTextOut "Sous-Signet 2", 15, 60
        
        PDFPrinter.PDFSetDrawColor = COLOR_CYAN
        PDFPrinter.PDFSetLineColor = COLOR_MAGENTA
        PDFPrinter.PDFSetLineStyle = pPDF_SOLID
        PDFPrinter.PDFSetLineWidth = 0.65
        PDFPrinter.PDFSetDrawMode = DRAW_DRAW
        PDFPrinter.PDFDrawRectangle 15, 80, 400, 50

        PDFPrinter.PDFSetLineColor = COLOR_BLEU
        PDFPrinter.PDFSetLineStyle = pPDF_SOLID
        PDFPrinter.PDFSetLineWidth = 0.85
        PDFPrinter.PDFSetDrawMode = DRAW_NORMAL
        PDFPrinter.PDFDrawRectangle 15, 150, 400, 50

        PDFPrinter.PDFImage App.Path & "\NeO78.jpg", 15, 220, 150, 100, "http://www.google.fr"
        
        PDFPrinter.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
        PDFPrinter.PDFSetTextColor = COLOR_NOIR
        PDFPrinter.PDFSetDrawColor = COLOR_JAUNE
        PDFPrinter.PDFSetLineColor = COLOR_NOIR
        PDFPrinter.PDFSetAlignement = ALIGN_LEFT
        PDFPrinter.PDFSetBorder = BORDER_ALL
        PDFPrinter.PDFSetLineWidth = 1
        PDFPrinter.PDFSetFill = True
        PDFPrinter.PDFCell "Test Alignement à Gauche", 15, 350, 150, 20
        
        PDFPrinter.PDFSetFont FONT_ARIAL, 8, FONT_ITALIC
        PDFPrinter.PDFSetTextColor = COLOR_NOIR
        PDFPrinter.PDFTextOut "Page " & PDFPrinter.PDFPageNumber, _
                         PDFPrinter.PDFGetPageWidth / 2 - PDFPrinter.PDFGetStringWidth("Page " & PDFPrinter.PDFPageNumber), _
                         PDFPrinter.PDFGetPageHeight - PDFPrinter.PDFTextHeight
                             
    PDFPrinter.PDFEndDoc
    
End Sub


