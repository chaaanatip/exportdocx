package main

import (
	"archive/zip"
	"bytes"
	"encoding/csv"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

// โครงสร้าง XML สำหรับ DOCX
type Document struct {
	XMLName  xml.Name `xml:"w:document"`
	Xmlns    string   `xml:"xmlns:w,attr"`
	XmlnsW14 string   `xml:"xmlns:w14,attr"`
	Body     Body     `xml:"w:body"`
}

type Body struct {
	XMLName xml.Name      `xml:"w:body"`
	Content []interface{} `xml:",any"`
	SectPr  SectPr        `xml:"w:sectPr"`
}

type SectPr struct {
	XMLName   xml.Name  `xml:"w:sectPr"`
	HeaderRef HeaderRef `xml:"w:headerReference"`
	FooterRef FooterRef `xml:"w:footerReference"`
	PgSz      PgSz      `xml:"w:pgSz"`
	PgMar     PgMar     `xml:"w:pgMar"`
}

type HeaderRef struct {
	XMLName xml.Name `xml:"w:headerReference"`
	Type    string   `xml:"w:type,attr"`
	Id      string   `xml:"r:id,attr"`
	Xmlns   string   `xml:"xmlns:r,attr"`
}

type FooterRef struct {
	XMLName xml.Name `xml:"w:footerReference"`
	Type    string   `xml:"w:type,attr"`
	Id      string   `xml:"r:id,attr"`
	Xmlns   string   `xml:"xmlns:r,attr"`
}

type PgSz struct {
	XMLName xml.Name `xml:"w:pgSz"`
	W       string   `xml:"w:w,attr"`
	H       string   `xml:"w:h,attr"`
}

type PgMar struct {
	XMLName xml.Name `xml:"w:pgMar"`
	Top     string   `xml:"w:top,attr"`
	Right   string   `xml:"w:right,attr"`
	Bottom  string   `xml:"w:bottom,attr"`
	Left    string   `xml:"w:left,attr"`
	Header  string   `xml:"w:header,attr"`
	Footer  string   `xml:"w:footer,attr"`
}

type Paragraph struct {
	XMLName      xml.Name      `xml:"w:p"`
	Props        *PPr          `xml:"w:pPr,omitempty"`
	Runs         []Run         `xml:"w:r"`
	Hyperlinks   []Hyperlink   `xml:"w:hyperlink,omitempty"`
	Bookmarks    []Bookmark    `xml:"w:bookmarkStart,omitempty"`
	BookmarkEnds []BookmarkEnd `xml:"w:bookmarkEnd,omitempty"`
}

type PPr struct {
	XMLName    xml.Name    `xml:"w:pPr"`
	Style      *PStyle     `xml:"w:pStyle,omitempty"`
	Jc         *Jc         `xml:"w:jc,omitempty"`
	Spacing    *Spacing    `xml:"w:spacing,omitempty"`
	NumPr      *NumPr      `xml:"w:numPr,omitempty"`
	OutlineLvl *OutlineLvl `xml:"w:outlineLvl,omitempty"`
	Ind        *Ind        `xml:"w:ind,omitempty"`
}

type Ind struct {
	XMLName   xml.Name `xml:"w:ind"`
	Left      string   `xml:"w:left,attr,omitempty"`
	Right     string   `xml:"w:right,attr,omitempty"`
	FirstLine string   `xml:"w:firstLine,attr,omitempty"`
	Hanging   string   `xml:"w:hanging,attr,omitempty"`
}

type PStyle struct {
	XMLName xml.Name `xml:"w:pStyle"`
	Val     string   `xml:"w:val,attr"`
}

type Spacing struct {
	XMLName xml.Name `xml:"w:spacing"`
	Before  string   `xml:"w:before,attr,omitempty"`
	After   string   `xml:"w:after,attr,omitempty"`
}

type NumPr struct {
	XMLName xml.Name `xml:"w:numPr"`
	ILvl    ILvl     `xml:"w:ilvl"`
	NumId   NumId    `xml:"w:numId"`
}

type ILvl struct {
	XMLName xml.Name `xml:"w:ilvl"`
	Val     string   `xml:"w:val,attr"`
}

type OutlineLvl struct {
	XMLName xml.Name `xml:"w:outlineLvl"`
	Val     string   `xml:"w:val,attr"`
}

type NumId struct {
	XMLName xml.Name `xml:"w:numId"`
	Val     string   `xml:"w:val,attr"`
}

type Jc struct {
	XMLName xml.Name `xml:"w:jc"`
	Val     string   `xml:"w:val,attr"`
}

type Run struct {
	XMLName xml.Name `xml:"w:r"`
	Props   *RPr     `xml:"w:rPr,omitempty"`
	Text    *Text    `xml:"w:t,omitempty"`
	Break   *Break   `xml:"w:br,omitempty"`
}

type RPr struct {
	XMLName   xml.Name   `xml:"w:rPr"`
	Bold      *Bold      `xml:"w:b,omitempty"`
	Italic    *Italic    `xml:"w:i,omitempty"`
	Underline *Underline `xml:"w:u,omitempty"`
	Size      *Size      `xml:"w:sz,omitempty"`
	Color     *Color     `xml:"w:color,omitempty"`
}

type Bold struct {
	XMLName xml.Name `xml:"w:b"`
}

type Italic struct {
	XMLName xml.Name `xml:"w:i"`
}

type Underline struct {
	Val string `xml:"w:val,attr"`
}

type Size struct {
	Val string `xml:"w:val,attr"`
}

type Color struct {
	Val string `xml:"w:val,attr"`
}

type Text struct {
	XMLName xml.Name `xml:"w:t"`
	Space   string   `xml:"xml:space,attr,omitempty"`
	Value   string   `xml:",chardata"`
}

type Break struct {
	XMLName xml.Name `xml:"w:br"`
	Type    string   `xml:"w:type,attr,omitempty"`
}

type Hyperlink struct {
	XMLName xml.Name `xml:"w:hyperlink"`
	Anchor  string   `xml:"w:anchor,attr,omitempty"`
	Tooltip string   `xml:"w:tooltip,attr,omitempty"`
	Runs    []Run    `xml:"w:r"`
}

type Bookmark struct {
	XMLName xml.Name `xml:"w:bookmarkStart"`
	Id      string   `xml:"w:id,attr"`
	Name    string   `xml:"w:name,attr"`
}

type BookmarkEnd struct {
	XMLName xml.Name `xml:"w:bookmarkEnd"`
	Id      string   `xml:"w:id,attr"`
}

// โครงสร้างข้อมูลสำหรับเก็บข้อมูลจาก CSV
type ChapterData struct {
	ID      string
	Chapter string
	Body    string
}

// CSS Class definitions และการแปลงเป็น Word properties
var cssClassMap = map[string]func(*PPr, *RPr){
	// Text Alignment
	"text-left":    func(pPr *PPr, rPr *RPr) { setJustification(pPr, "left") },
	"text-center":  func(pPr *PPr, rPr *RPr) { setJustification(pPr, "center") },
	"text-right":   func(pPr *PPr, rPr *RPr) { setJustification(pPr, "right") },
	"text-justify": func(pPr *PPr, rPr *RPr) { setJustification(pPr, "both") },
	"center":       func(pPr *PPr, rPr *RPr) { setJustification(pPr, "center") },
	"right":        func(pPr *PPr, rPr *RPr) { setJustification(pPr, "right") },
	"justify":      func(pPr *PPr, rPr *RPr) { setJustification(pPr, "both") },

	// Indentation (720 twips = 0.5 inch, 1440 twips = 1 inch)
	"indent-a":      func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "720", "") },
	"indent-b":      func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "1440", "") },
	"indent-c":      func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "2160", "") },
	"indent-1":      func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "720", "") },
	"indent-2":      func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "1440", "") },
	"indent-3":      func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "2160", "") },
	"indent-4":      func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "2880", "") },
	"indent-5":      func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "3600", "") },
	"indent-small":  func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "360", "") },
	"indent-medium": func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "720", "") },
	"indent-large":  func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "1440", "") },
	"indent-xl":     func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "2160", "") },

	// Hanging Indent
	"hanging-indent":   func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "720", "720") },
	"hanging-indent-1": func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "720", "720") },
	"hanging-indent-2": func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "1440", "720") },

	// Font Weight
	"bold":           func(pPr *PPr, rPr *RPr) { setBold(rPr, true) },
	"font-bold":      func(pPr *PPr, rPr *RPr) { setBold(rPr, true) },
	"font-semibold":  func(pPr *PPr, rPr *RPr) { setBold(rPr, true) },
	"font-extrabold": func(pPr *PPr, rPr *RPr) { setBold(rPr, true) },
	"font-black":     func(pPr *PPr, rPr *RPr) { setBold(rPr, true) },
	"font-normal":    func(pPr *PPr, rPr *RPr) { setBold(rPr, false) },
	"font-light":     func(pPr *PPr, rPr *RPr) { setBold(rPr, false) },

	// Font Style
	"italic":      func(pPr *PPr, rPr *RPr) { setItalic(rPr, true) },
	"font-italic": func(pPr *PPr, rPr *RPr) { setItalic(rPr, true) },
	"not-italic":  func(pPr *PPr, rPr *RPr) { setItalic(rPr, false) },

	// Text Decoration
	"underline":        func(pPr *PPr, rPr *RPr) { setUnderline(rPr, "single") },
	"underline-double": func(pPr *PPr, rPr *RPr) { setUnderline(rPr, "double") },
	"underline-dotted": func(pPr *PPr, rPr *RPr) { setUnderline(rPr, "dotted") },
	"underline-dashed": func(pPr *PPr, rPr *RPr) { setUnderline(rPr, "dash") },
	"no-underline":     func(pPr *PPr, rPr *RPr) { setUnderline(rPr, "none") },

	// Font Size (based on Tailwind CSS classes)
	"text-xs":   func(pPr *PPr, rPr *RPr) { setFontSize(rPr, "18") },  // 9pt
	"text-sm":   func(pPr *PPr, rPr *RPr) { setFontSize(rPr, "20") },  // 10pt
	"text-base": func(pPr *PPr, rPr *RPr) { setFontSize(rPr, "22") },  // 11pt
	"text-lg":   func(pPr *PPr, rPr *RPr) { setFontSize(rPr, "24") },  // 12pt
	"text-xl":   func(pPr *PPr, rPr *RPr) { setFontSize(rPr, "28") },  // 14pt
	"text-2xl":  func(pPr *PPr, rPr *RPr) { setFontSize(rPr, "32") },  // 16pt
	"text-3xl":  func(pPr *PPr, rPr *RPr) { setFontSize(rPr, "36") },  // 18pt
	"text-4xl":  func(pPr *PPr, rPr *RPr) { setFontSize(rPr, "48") },  // 24pt
	"text-5xl":  func(pPr *PPr, rPr *RPr) { setFontSize(rPr, "60") },  // 30pt
	"text-6xl":  func(pPr *PPr, rPr *RPr) { setFontSize(rPr, "72") },  // 36pt

	// Colors
	"text-black":   func(pPr *PPr, rPr *RPr) { setColor(rPr, "000000") },
	"text-white":   func(pPr *PPr, rPr *RPr) { setColor(rPr, "FFFFFF") },
	"text-red":     func(pPr *PPr, rPr *RPr) { setColor(rPr, "FF0000") },
	"text-green":   func(pPr *PPr, rPr *RPr) { setColor(rPr, "008000") },
	"text-blue":    func(pPr *PPr, rPr *RPr) { setColor(rPr, "0000FF") },
	"text-yellow":  func(pPr *PPr, rPr *RPr) { setColor(rPr, "FFFF00") },
	"text-purple":  func(pPr *PPr, rPr *RPr) { setColor(rPr, "800080") },
	"text-orange":  func(pPr *PPr, rPr *RPr) { setColor(rPr, "FFA500") },
	"text-gray":    func(pPr *PPr, rPr *RPr) { setColor(rPr, "808080") },
	"text-grey":    func(pPr *PPr, rPr *RPr) { setColor(rPr, "808080") },
	"red":          func(pPr *PPr, rPr *RPr) { setColor(rPr, "FF0000") },
	"green":        func(pPr *PPr, rPr *RPr) { setColor(rPr, "008000") },
	"blue":         func(pPr *PPr, rPr *RPr) { setColor(rPr, "0000FF") },
	"yellow":       func(pPr *PPr, rPr *RPr) { setColor(rPr, "FFFF00") },
	"purple":       func(pPr *PPr, rPr *RPr) { setColor(rPr, "800080") },
	"orange":       func(pPr *PPr, rPr *RPr) { setColor(rPr, "FFA500") },
	"gray":         func(pPr *PPr, rPr *RPr) { setColor(rPr, "808080") },
	"grey":         func(pPr *PPr, rPr *RPr) { setColor(rPr, "808080") },

	// Special styles
	"quote":      func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "720", ""); setItalic(rPr, true) },
	"blockquote": func(pPr *PPr, rPr *RPr) { setIndentation(pPr, "720", ""); setItalic(rPr, true) },
	"code":       func(pPr *PPr, rPr *RPr) { setColor(rPr, "800000") },
	"highlight":  func(pPr *PPr, rPr *RPr) { setColor(rPr, "FF6600"); setBold(rPr, true) },
	"emphasis":   func(pPr *PPr, rPr *RPr) { setItalic(rPr, true) },
	"strong":     func(pPr *PPr, rPr *RPr) { setBold(rPr, true) },
	"important":  func(pPr *PPr, rPr *RPr) { setBold(rPr, true); setColor(rPr, "FF0000") },
	"note":       func(pPr *PPr, rPr *RPr) { setColor(rPr, "0066CC"); setItalic(rPr, true) },
	"warning":    func(pPr *PPr, rPr *RPr) { setColor(rPr, "FF6600"); setBold(rPr, true) },
	"error":      func(pPr *PPr, rPr *RPr) { setColor(rPr, "FF0000"); setBold(rPr, true) },
	"success":    func(pPr *PPr, rPr *RPr) { setColor(rPr, "008000"); setBold(rPr, true) },

	// Spacing
	"spacing-tight":  func(pPr *PPr, rPr *RPr) { setSpacing(pPr, "", "40") },
	"spacing-normal": func(pPr *PPr, rPr *RPr) { setSpacing(pPr, "", "120") },
	"spacing-loose":  func(pPr *PPr, rPr *RPr) { setSpacing(pPr, "", "240") },
}

// Helper functions for setting Word properties
func setJustification(pPr *PPr, val string) {
	if pPr.Jc == nil {
		pPr.Jc = &Jc{}
	}
	pPr.Jc.Val = val
}

func setIndentation(pPr *PPr, left, hanging string) {
	if pPr.Ind == nil {
		pPr.Ind = &Ind{}
	}
	if left != "" {
		pPr.Ind.Left = left
	}
	if hanging != "" {
		pPr.Ind.Hanging = hanging
	}
}

func setBold(rPr *RPr, bold bool) {
	if bold {
		rPr.Bold = &Bold{}
	} else {
		rPr.Bold = nil
	}
}

func setItalic(rPr *RPr, italic bool) {
	if italic {
		rPr.Italic = &Italic{}
	} else {
		rPr.Italic = nil
	}
}

func setUnderline(rPr *RPr, val string) {
	if val == "none" {
		rPr.Underline = nil
	} else {
		rPr.Underline = &Underline{Val: val}
	}
}

func setFontSize(rPr *RPr, val string) {
	rPr.Size = &Size{Val: val}
}

func setColor(rPr *RPr, val string) {
	rPr.Color = &Color{Val: val}
}

func setSpacing(pPr *PPr, before, after string) {
	if pPr.Spacing == nil {
		pPr.Spacing = &Spacing{}
	}
	if before != "" {
		pPr.Spacing.Before = before
	}
	if after != "" {
		pPr.Spacing.After = after
	}
}

// ฟังก์ชันหลักสำหรับจัดการ CSS classes
func applyCSSClasses(classes []string, pPr *PPr, rPr *RPr) {
	for _, class := range classes {
		class = strings.TrimSpace(class)
		if class == "" {
			continue
		}

		// ตรวจหา CSS class ใน map
		if handler, exists := cssClassMap[class]; exists {
			handler(pPr, rPr)
		} else {
			// จัดการ dynamic classes
			handleDynamicCSSClass(class, pPr, rPr)
		}
	}
}

// จัดการ dynamic CSS classes
func handleDynamicCSSClass(class string, pPr *PPr, rPr *RPr) {
	// Font size with numbers: text-12, text-14, text-16, etc.
	sizeRegex := regexp.MustCompile(`^text-(\d+)$`)
	if matches := sizeRegex.FindStringSubmatch(class); len(matches) == 2 {
		if size, err := strconv.Atoi(matches[1]); err == nil {
			wordSize := size * 2 // Convert pt to half-points
			setFontSize(rPr, fmt.Sprintf("%d", wordSize))
		}
		return
	}

	// Margin classes: m-1, mt-2, mb-3, etc.
	marginRegex := regexp.MustCompile(`^m([tb]?)-(\d+)$`)
	if matches := marginRegex.FindStringSubmatch(class); len(matches) == 3 {
		direction := matches[1]
		value := matches[2]
		if spacing, err := strconv.Atoi(value); err == nil {
			twips := fmt.Sprintf("%d", spacing*120) // Convert to twips
			switch direction {
			case "t", "": // top or all
				setSpacing(pPr, twips, "")
			case "b": // bottom
				setSpacing(pPr, "", twips)
			}
		}
		return
	}

	// Padding/Indent classes: p-1, pl-2, etc.
	paddingRegex := regexp.MustCompile(`^p([l]?)-(\d+)$`)
	if matches := paddingRegex.FindStringSubmatch(class); len(matches) == 3 {
		direction := matches[1]
		value := matches[2]
		if indent, err := strconv.Atoi(value); err == nil {
			twips := fmt.Sprintf("%d", indent*360) // Convert to twips
			if direction == "l" || direction == "" {
				setIndentation(pPr, twips, "")
			}
		}
		return
	}
}

func main() {
	if len(os.Args) < 2 {
		fmt.Println("การใช้งาน: go run main.go <ไฟล์_csv>")
		fmt.Println("ตัวอย่าง: go run main.go data.csv")
		fmt.Println("CSV ต้องมี 3 คอลัมน์: id, chapter, body")
		os.Exit(1)
	}

	csvFile := os.Args[1]

	// ตรวจสอบว่าไฟล์ CSV มีอยู่จริง
	if _, err := os.Stat(csvFile); os.IsNotExist(err) {
		log.Fatalf("ไฟล์ %s ไม่พบ", csvFile)
	}

	// สร้างชื่อไฟล์ output
	docxFile := strings.TrimSuffix(csvFile, filepath.Ext(csvFile)) + ".docx"

	err := convertCSVToDocx(csvFile, docxFile)
	if err != nil {
		log.Fatalf("เกิดข้อผิดพลาดในการแปลงไฟล์: %v", err)
	}

	fmt.Printf("แปลงไฟล์สำเร็จ: %s -> %s\n", csvFile, docxFile)
}

func convertCSVToDocx(csvPath, docxPath string) error {
	// อ่านไฟล์ CSV
	chapters, err := readChapterCSV(csvPath)
	if err != nil {
		return fmt.Errorf("ไม่สามารถอ่านไฟล์ CSV: %w", err)
	}

	if len(chapters) == 0 {
		return fmt.Errorf("ไฟล์ CSV ว่างเปล่า")
	}

	// สร้าง DOCX
	return createDocxWithChapters(chapters, docxPath)
}

func readChapterCSV(filename string) ([]ChapterData, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	reader := csv.NewReader(file)
	records, err := reader.ReadAll()
	if err != nil {
		return nil, err
	}

	if len(records) < 2 { // ต้องมีอย่างน้อย header + 1 row
		return nil, fmt.Errorf("ไฟล์ CSV ต้องมีอย่างน้อย 2 แถว (header + data)")
	}

	// ตรวจสอบ header
	header := records[0]
	if len(header) < 3 {
		return nil, fmt.Errorf("CSV ต้องมี 3 คอลัมน์: id, chapter, body")
	}

	var chapters []ChapterData
	for i := 1; i < len(records); i++ {
		row := records[i]
		if len(row) < 3 {
			continue // ข้าม row ที่ไม่ครบ
		}

		chapters = append(chapters, ChapterData{
			ID:      strings.TrimSpace(row[0]),
			Chapter: strings.TrimSpace(row[1]),
			Body:    strings.TrimSpace(row[2]),
		})
	}

	return chapters, nil
}

func createDocxWithChapters(chapters []ChapterData, filename string) error {
	// สร้าง ZIP file สำหรับ DOCX
	docxFile, err := os.Create(filename)
	if err != nil {
		return err
	}
	defer docxFile.Close()

	zipWriter := zip.NewWriter(docxFile)
	defer zipWriter.Close()

	// สร้างไฟล์ที่จำเป็นใน DOCX
	if err := createContentTypes(zipWriter); err != nil {
		return err
	}

	if err := createRels(zipWriter); err != nil {
		return err
	}

	if err := createApp(zipWriter); err != nil {
		return err
	}

	if err := createCore(zipWriter); err != nil {
		return err
	}

	if err := createDocumentRels(zipWriter); err != nil {
		return err
	}

	if err := createStyles(zipWriter); err != nil {
		return err
	}

	if err := createSettings(zipWriter); err != nil {
		return err
	}

	if err := createHeader(zipWriter); err != nil {
		return err
	}

	if err := createFooter(zipWriter); err != nil {
		return err
	}

	if err := createWebSettings(zipWriter); err != nil {
		return err
	}

	if err := createDocumentWithChapters(zipWriter, chapters); err != nil {
		return err
	}

	return nil
}

func createDocumentWithChapters(zipWriter *zip.Writer, chapters []ChapterData) error {
	w, err := zipWriter.Create("word/document.xml")
	if err != nil {
		return err
	}

	// สร้างเอกสาร
	doc := Document{
		Xmlns:    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		XmlnsW14: "http://schemas.microsoft.com/office/word/2010/wordml",
		Body: Body{
			Content: []interface{}{},
			SectPr: SectPr{
				HeaderRef: HeaderRef{
					Type:  "default",
					Id:    "rId4",
					Xmlns: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
				},
				FooterRef: FooterRef{
					Type:  "default",
					Id:    "rId5",
					Xmlns: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
				},
				PgSz: PgSz{
					W: "11906",
					H: "16838",
				},
				PgMar: PgMar{
					Top:    "1440",
					Right:  "1440",
					Bottom: "1440",
					Left:   "1440",
					Header: "708",
					Footer: "708",
				},
			},
		},
	}

	// เพิ่มเนื้อหาแต่ละบท
	for i, chapter := range chapters {
		// เพิ่ม page break ก่อนตอนที่ 2 เป็นต้นไป
		if i > 0 {
			pageBreak := Paragraph{
				Runs: []Run{
					{
						Break: &Break{Type: "page"},
					},
				},
			}
			doc.Body.Content = append(doc.Body.Content, pageBreak)
		}

		bookmarkName := fmt.Sprintf("chapter_%d", i+1)

		// หัวข้อบทพร้อม bookmark
		chapterTitle := Paragraph{
			Props: &PPr{
				Style:      &PStyle{Val: "Heading2"},
				OutlineLvl: &OutlineLvl{Val: "1"},
				Spacing:    &Spacing{Before: "480", After: "240"},
			},
			Bookmarks: []Bookmark{
				{
					Id:   fmt.Sprintf("%d", i+1),
					Name: bookmarkName,
				},
			},
			BookmarkEnds: []BookmarkEnd{
				{
					Id: fmt.Sprintf("%d", i+1),
				},
			},
			Runs: []Run{
				{
					Props: &RPr{
						Bold: &Bold{},
						Size: &Size{Val: "24"},
					},
					Text: &Text{Value: fmt.Sprintf("ตอนที่ %d", i+1)},
				},
			},
		}
		doc.Body.Content = append(doc.Body.Content, chapterTitle)

		// แปลง HTML เป็น paragraphs
		bodyParagraphs := parseHTMLToWordParagraphs(chapter.Body)
		for _, para := range bodyParagraphs {
			doc.Body.Content = append(doc.Body.Content, para)
		}
	}

	// เขียน XML
	xmlHeader := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"

	var buf bytes.Buffer
	encoder := xml.NewEncoder(&buf)
	encoder.Indent("", "  ")

	if err := encoder.Encode(doc); err != nil {
		return err
	}

	_, err = w.Write([]byte(xmlHeader))
	if err != nil {
		return err
	}

	_, err = io.Copy(w, &buf)
	return err
}

func parseHTMLToWordParagraphs(htmlContent string) []Paragraph {
	var paragraphs []Paragraph

	// ลบ HTML comments
	htmlContent = regexp.MustCompile(`<!--.*?-->`).ReplaceAllString(htmlContent, "")

	// แยก paragraphs โดยใช้ <p> tags
	pRegex := regexp.MustCompile(`<p[^>]*>(.*?)</p>`)
	pMatches := pRegex.FindAllStringSubmatch(htmlContent, -1)

	if len(pMatches) == 0 {
		// ถ้าไม่มี <p> tags ให้แยกด้วย <br> หรือ newlines
		lines := regexp.MustCompile(`<br\s*/?>\s*|<br[^>]*>\s*|\n\n+`).Split(htmlContent, -1)
		for _, line := range lines {
			line = strings.TrimSpace(line)
			if line != "" {
				para := parseHTMLLineToParagraph(line)
				paragraphs = append(paragraphs, para)
			}
		}
	} else {
		// มี <p> tags
		for _, match := range pMatches {
			content := strings.TrimSpace(match[1])
			if content != "" {
				para := parseHTMLLineToParagraph(content)
				paragraphs = append(paragraphs, para)
			}
		}
	}

	// ถ้าไม่มี paragraphs เลย ให้สร้าง paragraph ธรรมดา
	if len(paragraphs) == 0 {
		para := parseHTMLLineToParagraph(htmlContent)
		paragraphs = append(paragraphs, para)
	}

	return paragraphs
}

func parseHTMLLineToParagraph(htmlLine string) Paragraph {
	para := Paragraph{
		Props: &PPr{
			Spacing: &Spacing{After: "80"},
		},
		Runs: []Run{},
	}

	// Initialize run properties for class handling
	runProps := &RPr{}

	// ตรวจหา CSS classes ใน HTML tags
	classRegex := regexp.MustCompile(`class=["']([^"']*)["']`)
	classMatches := classRegex.FindAllStringSubmatch(htmlLine, -1)

	for _, match := range classMatches {
		if len(match) > 1 {
			classes := strings.Split(match[1], " ")
			applyCSSClasses(classes, para.Props, runProps)
		}
	}

	// ตรวจหา inline styles
	styleRegex := regexp.MustCompile(`style=["']([^"']*)["']`)
	styleMatches := styleRegex.FindAllStringSubmatch(htmlLine, -1)

	for _, match := range styleMatches {
		if len(match) > 1 {
			parseInlineStyles(match[1], para.Props, runProps)
		}
	}

	// แปลง HTML tags เป็น Word formatting
	runs := parseHTMLToRunsWithGlobalProps(htmlLine, runProps)
	para.Runs = runs

	return para
}

// Parse inline CSS styles
func parseInlineStyles(styles string, pPr *PPr, rPr *RPr) {
	// Text alignment
	if regexp.MustCompile(`text-align:\s*center`).MatchString(styles) {
		setJustification(pPr, "center")
	} else if regexp.MustCompile(`text-align:\s*right`).MatchString(styles) {
		setJustification(pPr, "right")
	} else if regexp.MustCompile(`text-align:\s*justify`).MatchString(styles) {
		setJustification(pPr, "both")
	} else if regexp.MustCompile(`text-align:\s*left`).MatchString(styles) {
		setJustification(pPr, "left")
	}

	// Font weight
	if regexp.MustCompile(`font-weight:\s*bold`).MatchString(styles) {
		setBold(rPr, true)
	}

	// Font style
	if regexp.MustCompile(`font-style:\s*italic`).MatchString(styles) {
		setItalic(rPr, true)
	}

	// Text decoration
	if regexp.MustCompile(`text-decoration:\s*underline`).MatchString(styles) {
		setUnderline(rPr, "single")
	}

	// Margin-left for indentation
	marginRegex := regexp.MustCompile(`margin-left:\s*(\d+(?:\.\d+)?)(?:px|em|pt)`)
	marginMatches := marginRegex.FindStringSubmatch(styles)
	if len(marginMatches) > 1 {
		if margin, err := strconv.ParseFloat(marginMatches[1], 64); err == nil {
			twips := int(margin * 15) // Convert pixels to twips
			setIndentation(pPr, fmt.Sprintf("%d", twips), "")
		}
	}

	// Padding-left for indentation
	paddingRegex := regexp.MustCompile(`padding-left:\s*(\d+(?:\.\d+)?)(?:px|em|pt)`)
	paddingMatches := paddingRegex.FindStringSubmatch(styles)
	if len(paddingMatches) > 1 {
		if padding, err := strconv.ParseFloat(paddingMatches[1], 64); err == nil {
			twips := int(padding * 15)
			if pPr.Ind == nil {
				pPr.Ind = &Ind{}
			}
			// ถ้ามี margin-left อยู่แล้ว ให้รวมกับ padding
			if pPr.Ind.Left != "" {
				if currentLeft, err := strconv.Atoi(pPr.Ind.Left); err == nil {
					twips += currentLeft
				}
			}
			pPr.Ind.Left = fmt.Sprintf("%d", twips)
		}
	}

	// Color
	colorPatterns := []string{
		`color:\s*#([0-9a-fA-F]{6})`,                                        // #FF0000
		`color:\s*#([0-9a-fA-F]{3})`,                                        // #F00
		`color:\s*rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)`,           // rgb(255,0,0)
		`color:\s*hsl\s*\(\s*(\d+)\s*,\s*(\d+)%\s*,\s*(\d+)%\s*\)`,         // hsl(30,100%,50%)
		`color:\s*([a-zA-Z]+)`,                                              // red, blue, etc.
	}

	for i, pattern := range colorPatterns {
		colorRegex := regexp.MustCompile(pattern)
		colorMatches := colorRegex.FindStringSubmatch(styles)

		if len(colorMatches) > 1 {
			var colorHex string

			switch i {
			case 0: // #FF0000
				colorHex = strings.ToUpper(colorMatches[1])
			case 1: // #F00
				hex3 := colorMatches[1]
				if len(hex3) == 3 {
					colorHex = strings.ToUpper(string(hex3[0]) + string(hex3[0]) + string(hex3[1]) + string(hex3[1]) + string(hex3[2]) + string(hex3[2]))
				}
			case 2: // rgb(255,0,0)
				if len(colorMatches) >= 4 {
					r, _ := strconv.Atoi(colorMatches[1])
					g, _ := strconv.Atoi(colorMatches[2])
					b, _ := strconv.Atoi(colorMatches[3])
					colorHex = fmt.Sprintf("%02X%02X%02X", r, g, b)
				}
			case 3: // hsl(30,100%,50%)
				if len(colorMatches) >= 4 {
					h, _ := strconv.Atoi(colorMatches[1])
					s, _ := strconv.Atoi(colorMatches[2])
					l, _ := strconv.Atoi(colorMatches[3])
					colorHex = convertHSLToHex(h, s, l)
				}
			case 4: // named colors
				colorName := strings.ToLower(colorMatches[1])
				colorHex = convertNamedColorToHex(colorName)
			}

			if colorHex != "" {
				setColor(rPr, colorHex)
				break
			}
		}
	}

	// Font size
	sizeRegex := regexp.MustCompile(`font-size:\s*(\d+)px`)
	sizeMatches := sizeRegex.FindStringSubmatch(styles)
	if len(sizeMatches) > 1 {
		if px, err := strconv.Atoi(sizeMatches[1]); err == nil {
			points := float64(px) * 0.75
			wordSize := int(points * 2)
			setFontSize(rPr, fmt.Sprintf("%d", wordSize))
		}
	}
}

func parseHTMLToRunsWithGlobalProps(htmlContent string, globalProps *RPr) []Run {
	var runs []Run

	// แปลง HTML entities ก่อนทำอย่างอื่น
	content := decodeHTMLEntities(htmlContent)

	// จัดการ <a> tags (links) ก่อน
	linkRegex := regexp.MustCompile(`<a[^>]*>(.*?)</a>`)
	for linkRegex.MatchString(content) {
		matches := linkRegex.FindAllStringSubmatch(content, -1)
		for _, match := range matches {
			linkText := strings.TrimSpace(match[1])
			if linkText != "" {
				linkText = regexp.MustCompile(`<[^>]+>`).ReplaceAllString(linkText, "")
				linkText = decodeHTMLEntities(linkText)

				run := Run{
					Props: mergeRunProps(globalProps, &RPr{
						Color:     &Color{Val: "0563C1"},
						Underline: &Underline{Val: "single"},
					}),
					Text: &Text{Value: linkText, Space: "preserve"},
				}
				runs = append(runs, run)
			}
		}
		content = linkRegex.ReplaceAllString(content, "")
	}

	// จัดการ <span> tags with styles และ classes
	spanRegex := regexp.MustCompile(`<span[^>]*>(.*?)</span>`)
	for spanRegex.MatchString(content) {
		matches := spanRegex.FindAllStringSubmatch(content, -1)
		for _, match := range matches {
			spanText := strings.TrimSpace(match[1])
			if spanText != "" {
				styleProps := extractStyleAndClassFromSpan(match[0])
				spanText = regexp.MustCompile(`<[^>]+>`).ReplaceAllString(spanText, "")
				spanText = decodeHTMLEntities(spanText)

				run := Run{
					Props: mergeRunProps(globalProps, styleProps),
					Text:  &Text{Value: spanText, Space: "preserve"},
				}
				runs = append(runs, run)
			}
		}
		content = spanRegex.ReplaceAllString(content, "")
	}

	// จัดการ <strong> และ <b> tags
	strongRegex := regexp.MustCompile(`<(strong|b)[^>]*>(.*?)</(strong|b)>`)
	for strongRegex.MatchString(content) {
		matches := strongRegex.FindAllStringSubmatch(content, -1)
		for _, match := range matches {
			boldText := strings.TrimSpace(match[2])
			if boldText != "" {
				boldText = regexp.MustCompile(`<[^>]+>`).ReplaceAllString(boldText, "")
				boldText = decodeHTMLEntities(boldText)

				run := Run{
					Props: mergeRunProps(globalProps, &RPr{
						Bold: &Bold{},
					}),
					Text: &Text{Value: boldText, Space: "preserve"},
				}
				runs = append(runs, run)
			}
		}
		content = strongRegex.ReplaceAllString(content, "")
	}

	// จัดการ <em> และ <i> tags
	emRegex := regexp.MustCompile(`<(em|i)[^>]*>(.*?)</(em|i)>`)
	for emRegex.MatchString(content) {
		matches := emRegex.FindAllStringSubmatch(content, -1)
		for _, match := range matches {
			italicText := strings.TrimSpace(match[2])
			if italicText != "" {
				italicText = regexp.MustCompile(`<[^>]+>`).ReplaceAllString(italicText, "")
				italicText = decodeHTMLEntities(italicText)

				run := Run{
					Props: mergeRunProps(globalProps, &RPr{
						Italic: &Italic{},
					}),
					Text: &Text{Value: italicText, Space: "preserve"},
				}
				runs = append(runs, run)
			}
		}
		content = emRegex.ReplaceAllString(content, "")
	}

	// จัดการ <u> tags
	uRegex := regexp.MustCompile(`<u[^>]*>(.*?)</u>`)
	for uRegex.MatchString(content) {
		matches := uRegex.FindAllStringSubmatch(content, -1)
		for _, match := range matches {
			underlineText := strings.TrimSpace(match[1])
			if underlineText != "" {
				underlineText = regexp.MustCompile(`<[^>]+>`).ReplaceAllString(underlineText, "")
				underlineText = decodeHTMLEntities(underlineText)

				run := Run{
					Props: mergeRunProps(globalProps, &RPr{
						Underline: &Underline{Val: "single"},
					}),
					Text: &Text{Value: underlineText, Space: "preserve"},
				}
				runs = append(runs, run)
			}
		}
		content = uRegex.ReplaceAllString(content, "")
	}

	// ลบ HTML tags ที่เหลือทั้งหมด
	content = regexp.MustCompile(`<[^>]+>`).ReplaceAllString(content, "")

	// แปลง HTML entities อีกครั้งสำหรับข้อความที่เหลือ
	content = decodeHTMLEntities(content)
	content = strings.TrimSpace(content)

	// เพิ่ม text ธรรมดาที่เหลือ
	if content != "" {
		run := Run{
			Props: globalProps,
			Text:  &Text{Value: content, Space: "preserve"},
		}
		runs = append(runs, run)
	}

	// ถ้าไม่มี runs เลย ให้สร้าง run ว่าง
	if len(runs) == 0 {
		runs = append(runs, Run{
			Props: globalProps,
			Text:  &Text{Value: "", Space: "preserve"},
		})
	}

	return runs
}

// ฟังก์ชันรวม run properties
func mergeRunProps(base, overlay *RPr) *RPr {
	if base == nil && overlay == nil {
		return nil
	}
	if base == nil {
		return overlay
	}
	if overlay == nil {
		return base
	}

	merged := &RPr{}

	// Bold
	if overlay.Bold != nil {
		merged.Bold = overlay.Bold
	} else if base.Bold != nil {
		merged.Bold = base.Bold
	}

	// Italic
	if overlay.Italic != nil {
		merged.Italic = overlay.Italic
	} else if base.Italic != nil {
		merged.Italic = base.Italic
	}

	// Underline
	if overlay.Underline != nil {
		merged.Underline = overlay.Underline
	} else if base.Underline != nil {
		merged.Underline = base.Underline
	}

	// Size
	if overlay.Size != nil {
		merged.Size = overlay.Size
	} else if base.Size != nil {
		merged.Size = base.Size
	}

	// Color
	if overlay.Color != nil {
		merged.Color = overlay.Color
	} else if base.Color != nil {
		merged.Color = base.Color
	}

	return merged
}

// ฟังก์ชันที่รองรับทั้ง style และ class
func extractStyleAndClassFromSpan(spanTag string) *RPr {
	props := &RPr{}

	// ตรวจหา class attribute ก่อน
	classRegex := regexp.MustCompile(`class=["']([^"']*)["']`)
	classMatches := classRegex.FindStringSubmatch(spanTag)

	if len(classMatches) > 1 {
		classes := strings.Split(classMatches[1], " ")
		// สร้าง dummy PPr เพื่อใช้กับ applyCSSClasses
		dummyPPr := &PPr{}
		applyCSSClasses(classes, dummyPPr, props)
	}

	// ดึง style attribute
	styleRegex := regexp.MustCompile(`style=["']([^"']*)["']`)
	matches := styleRegex.FindStringSubmatch(spanTag)

	if len(matches) > 1 {
		// สร้าง dummy PPr เพื่อใช้กับ parseInlineStyles
		dummyPPr := &PPr{}
		parseInlineStyles(matches[1], dummyPPr, props)
	}

	return props
}

func convertHSLToHex(h, s, l int) string {
	// Simplified HSL to RGB conversion
	if s == 0 {
		gray := l * 255 / 100
		return fmt.Sprintf("%02X%02X%02X", gray, gray, gray)
	}

	// Common color mappings for typical HSL values
	colorMap := map[string]string{
		"30_100_50":  "FF6600", // orange
		"0_100_50":   "FF0000", // red
		"120_100_50": "00FF00", // green
		"240_100_50": "0000FF", // blue
		"60_100_50":  "FFFF00", // yellow
		"300_100_50": "FF00FF", // magenta
		"180_100_50": "00FFFF", // cyan
	}

	key := fmt.Sprintf("%d_%d_%d", h, s, l)
	if hex, exists := colorMap[key]; exists {
		return hex
	}

	return "000000" // Default to black
}

func convertNamedColorToHex(colorName string) string {
	colorMap := map[string]string{
		"black":   "000000",
		"white":   "FFFFFF",
		"red":     "FF0000",
		"green":   "008000",
		"blue":    "0000FF",
		"yellow":  "FFFF00",
		"cyan":    "00FFFF",
		"magenta": "FF00FF",
		"orange":  "FFA500",
		"purple":  "800080",
		"brown":   "A52A2A",
		"pink":    "FFC0CB",
		"gray":    "808080",
		"grey":    "808080",
		"lime":    "00FF00",
		"navy":    "000080",
		"maroon":  "800000",
		"olive":   "808000",
		"teal":    "008080",
		"silver":  "C0C0C0",
		"gold":    "FFD700",
		"violet":  "EE82EE",
		"indigo":  "4B0082",
		"crimson": "DC143C",
		"coral":   "FF7F50",
	}

	if hex, exists := colorMap[colorName]; exists {
		return hex
	}

	return "" // ไม่รู้จักสี
}

func decodeHTMLEntities(text string) string {
	// ทำการแปลงหลายรอบเพื่อให้แน่ใจว่าได้หมด
	for i := 0; i < 3; i++ {
		oldText := text
		text = performEntityDecoding(text)
		if text == oldText {
			break // ถ้าไม่มีการเปลี่ยนแปลงแล้วให้หยุด
		}
	}
	return text
}

func performEntityDecoding(text string) string {
	// Basic HTML entities
	text = strings.ReplaceAll(text, "&amp;", "&")
	text = strings.ReplaceAll(text, "&lt;", "<")
	text = strings.ReplaceAll(text, "&gt;", ">")
	text = strings.ReplaceAll(text, "&quot;", "\"")
	text = strings.ReplaceAll(text, "&#39;", "'")
	text = strings.ReplaceAll(text, "&apos;", "'")
	text = strings.ReplaceAll(text, "&nbsp;", " ")

	// Quotation marks (ทั้ง named และ numeric)
	text = strings.ReplaceAll(text, "&lsquo;", "'")      // left single quote
	text = strings.ReplaceAll(text, "&rsquo;", "'")      // right single quote
	text = strings.ReplaceAll(text, "&ldquo;", "\u201C") // left double quote
	text = strings.ReplaceAll(text, "&rdquo;", "\u201D") // right double quote
	text = strings.ReplaceAll(text, "&#8216;", "'")      // left single quote
	text = strings.ReplaceAll(text, "&#8217;", "'")      // right single quote
	text = strings.ReplaceAll(text, "&#8218;", "\u201A") // single low quote
	text = strings.ReplaceAll(text, "&#8219;", "\u201B") // single high reversed quote
	text = strings.ReplaceAll(text, "&#8220;", "\u201C") // left double quote
	text = strings.ReplaceAll(text, "&#8221;", "\u201D") // right double quote
	text = strings.ReplaceAll(text, "&#8222;", "\u201E") // double low quote

	// Dashes
	text = strings.ReplaceAll(text, "&mdash;", "\u2014") // em dash
	text = strings.ReplaceAll(text, "&ndash;", "\u2013") // en dash
	text = strings.ReplaceAll(text, "&#8212;", "\u2014") // em dash
	text = strings.ReplaceAll(text, "&#8211;", "\u2013") // en dash
	text = strings.ReplaceAll(text, "&#8208;", "-")      // hyphen
	text = strings.ReplaceAll(text, "&#8209;", "\u2011") // non-breaking hyphen
	text = strings.ReplaceAll(text, "&#8210;", "\u2012") // figure dash

	// Other punctuation
	text = strings.ReplaceAll(text, "&hellip;", "\u2026") // ellipsis
	text = strings.ReplaceAll(text, "&#8230;", "\u2026")  // ellipsis
	text = strings.ReplaceAll(text, "&#8226;", "\u2022")  // bullet
	text = strings.ReplaceAll(text, "&bull;", "\u2022")   // bullet
	text = strings.ReplaceAll(text, "&#8240;", "\u2030")  // per mille
	text = strings.ReplaceAll(text, "&permil;", "\u2030") // per mille

	// Copyright and trademark
	text = strings.ReplaceAll(text, "&copy;", "\u00A9")   // copyright
	text = strings.ReplaceAll(text, "&#169;", "\u00A9")   // copyright
	text = strings.ReplaceAll(text, "&reg;", "\u00AE")    // registered
	text = strings.ReplaceAll(text, "&#174;", "\u00AE")   // registered
	text = strings.ReplaceAll(text, "&trade;", "\u2122")  // trademark
	text = strings.ReplaceAll(text, "&#8482;", "\u2122") // trademark
	text = strings.ReplaceAll(text, "&#8471;", "\u2117") // sound recording copyright

	// Math symbols
	text = strings.ReplaceAll(text, "&plusmn;", "\u00B1") // plus-minus
	text = strings.ReplaceAll(text, "&#177;", "\u00B1")   // plus-minus
	text = strings.ReplaceAll(text, "&times;", "\u00D7")  // multiplication
	text = strings.ReplaceAll(text, "&#215;", "\u00D7")   // multiplication
	text = strings.ReplaceAll(text, "&divide;", "\u00F7") // division
	text = strings.ReplaceAll(text, "&#247;", "\u00F7")   // division
	text = strings.ReplaceAll(text, "&ne;", "\u2260")     // not equal
	text = strings.ReplaceAll(text, "&#8800;", "\u2260")  // not equal

	// Currency
	text = strings.ReplaceAll(text, "&cent;", "\u00A2")
	text = strings.ReplaceAll(text, "&#162;", "\u00A2")
	text = strings.ReplaceAll(text, "&pound;", "\u00A3")
	text = strings.ReplaceAll(text, "&#163;", "\u00A3")
	text = strings.ReplaceAll(text, "&yen;", "\u00A5")
	text = strings.ReplaceAll(text, "&#165;", "\u00A5")
	text = strings.ReplaceAll(text, "&euro;", "\u20AC")
	text = strings.ReplaceAll(text, "&#8364;", "\u20AC")

	// Spaces
	text = strings.ReplaceAll(text, "&#160;", " ")  // non-breaking space
	text = strings.ReplaceAll(text, "&#8201;", " ") // thin space
	text = strings.ReplaceAll(text, "&#8194;", " ") // en space
	text = strings.ReplaceAll(text, "&#8195;", " ") // em space

	// Decode numeric character references (decimal)
	numericRegex := regexp.MustCompile(`&#(\d+);`)
	text = numericRegex.ReplaceAllStringFunc(text, func(match string) string {
		numStr := strings.TrimPrefix(strings.TrimSuffix(match, ";"), "&#")
		if num, err := strconv.Atoi(numStr); err == nil && num > 0 && num <= 1114111 {
			return string(rune(num))
		}
		return match
	})

	// Decode hexadecimal character references
	hexRegex := regexp.MustCompile(`&#[xX]([0-9a-fA-F]+);`)
	text = hexRegex.ReplaceAllStringFunc(text, func(match string) string {
		hexStr := strings.TrimPrefix(strings.TrimSuffix(match, ";"), "&#")
		hexStr = strings.TrimPrefix(strings.ToLower(hexStr), "x")
		if num, err := strconv.ParseInt(hexStr, 16, 32); err == nil && num > 0 && num <= 1114111 {
			return string(rune(num))
		}
		return match
	})

	return text
}

func createContentTypes(zipWriter *zip.Writer) error {
	w, err := zipWriter.Create("[Content_Types].xml")
	if err != nil {
		return err
	}

	content := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
    <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
    <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
    <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`

	_, err = w.Write([]byte(content))
	return err
}

func createRels(zipWriter *zip.Writer) error {
	w, err := zipWriter.Create("_rels/.rels")
	if err != nil {
		return err
	}

	content := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`

	_, err = w.Write([]byte(content))
	return err
}

func createApp(zipWriter *zip.Writer) error {
	w, err := zipWriter.Create("docProps/app.xml")
	if err != nil {
		return err
	}

	content := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
    <Application>CSV Chapter to DOCX Converter</Application>
    <DocSecurity>0</DocSecurity>
    <ScaleCrop>false</ScaleCrop>
    <SharedDoc>false</SharedDoc>
    <HyperlinksChanged>false</HyperlinksChanged>
    <AppVersion>1.0</AppVersion>
</Properties>`

	_, err = w.Write([]byte(content))
	return err
}

func createCore(zipWriter *zip.Writer) error {
	w, err := zipWriter.Create("docProps/core.xml")
	if err != nil {
		return err
	}

	content := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:title>Chapter Document</dc:title>
    <dc:creator>CSV Chapter to DOCX Converter</dc:creator>
    <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>`

	_, err = w.Write([]byte(content))
	return err
}

func createDocumentRels(zipWriter *zip.Writer) error {
	w, err := zipWriter.Create("word/_rels/document.xml.rels")
	if err != nil {
		return err
	}

	content := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
    <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`

	_, err = w.Write([]byte(content))
	return err
}

func createSettings(zipWriter *zip.Writer) error {
	w, err := zipWriter.Create("word/settings.xml")
	if err != nil {
		return err
	}

	content := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:defaultTabStop w:val="708"/>
    <w:characterSpacingControl w:val="doNotCompress"/>
    <w:compat>
        <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
    </w:compat>
    <w:rsids>
        <w:rsidRoot w:val="00000000"/>
    </w:rsids>
    <w:mathPr>
        <w:mathFont w:val="Cambria Math"/>
        <w:brkBin w:val="before"/>
        <w:brkBinSub w:val="--"/>
        <w:smallFrac w:val="0"/>
        <w:dispDef/>
        <w:lMargin w:val="0"/>
        <w:rMargin w:val="0"/>
        <w:defJc w:val="centerGroup"/>
        <w:wrapIndent w:val="1440"/>
        <w:intLim w:val="subSup"/>
        <w:naryLim w:val="undOvr"/>
    </w:mathPr>
    <w:view w:val="outline"/>
    <w:zoom w:percent="100"/>
    <w:doNotDisplayPageBoundaries/>
    <w:displayBackgroundShape w:val="false"/>
</w:settings>`

	_, err = w.Write([]byte(content))
	return err
}

func createHeader(zipWriter *zip.Writer) error {
	w, err := zipWriter.Create("word/header1.xml")
	if err != nil {
		return err
	}

	content := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:p>
        <w:pPr>
            <w:jc w:val="center"/>
            <w:spacing w:after="120"/>
        </w:pPr>
        <w:r>
            <w:rPr>
                <w:b/>
                <w:sz w:val="20"/>
            </w:rPr>
            <w:t>เอกสารจากไฟล์ CSV</w:t>
        </w:r>
    </w:p>
</w:hdr>`

	_, err = w.Write([]byte(content))
	return err
}

func createFooter(zipWriter *zip.Writer) error {
	w, err := zipWriter.Create("word/footer1.xml")
	if err != nil {
		return err
	}

	content := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:p>
        <w:pPr>
            <w:jc w:val="center"/>
        </w:pPr>
        <w:r>
            <w:t>หน้า </w:t>
        </w:r>
        <w:fldSimple w:instr=" PAGE ">
            <w:r>
                <w:t>1</w:t>
            </w:r>
        </w:fldSimple>
        <w:r>
            <w:t> จาก </w:t>
        </w:r>
        <w:fldSimple w:instr=" NUMPAGES ">
            <w:r>
                <w:t>1</w:t>
            </w:r>
        </w:fldSimple>
    </w:p>
</w:ftr>`

	_, err = w.Write([]byte(content))
	return err
}

func createWebSettings(zipWriter *zip.Writer) error {
	w, err := zipWriter.Create("word/webSettings.xml")
	if err != nil {
		return err
	}

	content := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:optimizeForBrowser/>
    <w:allowPNG/>
    <w:doNotRelyOnCSS/>
    <w:doNotSaveAsSingleFile/>
</w:webSettings>`

	_, err = w.Write([]byte(content))
	return err
}

func createStyles(zipWriter *zip.Writer) error {
	w, err := zipWriter.Create("word/styles.xml")
	if err != nil {
		return err
	}

	content := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:docDefaults>
        <w:rPrDefault>
            <w:rPr>
                <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
                <w:sz w:val="22"/>
            </w:rPr>
        </w:rPrDefault>
        <w:pPrDefault>
            <w:pPr>
                <w:spacing w:after="120" w:line="276" w:lineRule="auto"/>
            </w:pPr>
        </w:pPrDefault>
    </w:docDefaults>
    
    <w:style w:type="paragraph" w:styleId="Normal">
        <w:name w:val="Normal"/>
        <w:qFormat/>
        <w:pPr>
            <w:spacing w:after="120"/>
        </w:pPr>
    </w:style>
    
    <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:qFormat/>
        <w:pPr>
            <w:keepNext/>
            <w:keepLines/>
            <w:spacing w:before="240" w:after="120"/>
            <w:outlineLvl w:val="0"/>
        </w:pPr>
        <w:rPr>
            <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
            <w:b/>
            <w:color w:val="365F91"/>
            <w:sz w:val="28"/>
        </w:rPr>
    </w:style>
    
    <w:style w:type="paragraph" w:styleId="Heading2">
        <w:name w:val="heading 2"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:qFormat/>
        <w:pPr>
            <w:keepNext/>
            <w:keepLines/>
            <w:spacing w:before="200" w:after="60"/>
            <w:outlineLvl w:val="1"/>
        </w:pPr>
        <w:rPr>
            <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
            <w:b/>
            <w:color w:val="4F81BD"/>
            <w:sz w:val="24"/>
        </w:rPr>
    </w:style>
    
</w:styles>`

	_, err = w.Write([]byte(content))
	return err
}