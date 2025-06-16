package main

import (
	"archive/zip"
	"bytes"
	"encoding/csv"
	"encoding/xml"
	"fmt"
	"html"
	"io"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

// โครงสร้างข้อมูลสำหรับ CSV

type Hyperlink struct {
    XMLName xml.Name `xml:"w:hyperlink"`
    Id      string   `xml:"r:id,attr"`
    Runs    []Run    `xml:"w:r"`
}

type ChapterData struct {
	ID      string
	Chapter string
	Body    string
}

// DOCX XML Structures
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
	XMLName xml.Name `xml:"w:sectPr"`
	PgSz    PgSz     `xml:"w:pgSz"`
	PgMar   PgMar    `xml:"w:pgMar"`
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
}

type Paragraph struct {
	XMLName xml.Name `xml:"w:p"`
	Props   *PPr     `xml:"w:pPr,omitempty"`
	Runs    []Run    `xml:"w:r"`
}

type PPr struct {
	XMLName   xml.Name  `xml:"w:pPr"`
	Jc        *Jc       `xml:"w:jc,omitempty"`
	Spacing   *Spacing  `xml:"w:spacing,omitempty"`
	Ind       *Ind      `xml:"w:ind,omitempty"`
	PStyle    *PStyle   `xml:"w:pStyle,omitempty"`
	OutlineLvl *OutlineLvl `xml:"w:outlineLvl,omitempty"`
}

type Run struct {
	XMLName xml.Name `xml:"w:r"`
	Props   *RPr     `xml:"w:rPr,omitempty"`
	Text    *Text    `xml:"w:t,omitempty"`
	Break   *Break   `xml:"w:br,omitempty"`
}

type RPr struct {
	XMLName xml.Name `xml:"w:rPr"`
	Bold    *Bold    `xml:"w:b,omitempty"`
	Italic  *Italic  `xml:"w:i,omitempty"`
	Color   *Color   `xml:"w:color,omitempty"`
	Size    *Size    `xml:"w:sz,omitempty"`
}

type Text struct {
	XMLName xml.Name `xml:"w:t"`
	Space   string   `xml:"xml:space,attr,omitempty"`
	Value   string   `xml:",chardata"`
}

type Bold struct {
	XMLName xml.Name `xml:"w:b"`
}

type Italic struct {
	XMLName xml.Name `xml:"w:i"`
}

type Color struct {
	XMLName xml.Name `xml:"w:color"`
	Val     string   `xml:"w:val,attr"`
}

type Size struct {
	XMLName xml.Name `xml:"w:sz"`
	Val     string   `xml:"w:val,attr"`
}

type Break struct {
	XMLName xml.Name `xml:"w:br"`
	Type    string   `xml:"w:type,attr,omitempty"`
}

type Jc struct {
	XMLName xml.Name `xml:"w:jc"`
	Val     string   `xml:"w:val,attr"`
}

type Spacing struct {
	XMLName xml.Name `xml:"w:spacing"`
	Before  string   `xml:"w:before,attr,omitempty"`
	After   string   `xml:"w:after,attr,omitempty"`
}

type Ind struct {
	XMLName   xml.Name `xml:"w:ind"`
	Left      string   `xml:"w:left,attr,omitempty"`
	FirstLine string   `xml:"w:firstLine,attr,omitempty"`
	Hanging   string   `xml:"w:hanging,attr,omitempty"`
}

type PStyle struct {
	XMLName xml.Name `xml:"w:pStyle"`
	Val     string   `xml:"w:val,attr"`
}

type OutlineLvl struct {
	XMLName xml.Name `xml:"w:outlineLvl"`
	Val     string   `xml:"w:val,attr"`
}

func main() {
	if len(os.Args) < 2 {
		fmt.Println("การใช้งาน: go run main.go <ไฟล์_csv>")
		fmt.Println("ตัวอย่าง: go run main.go data.csv")
		os.Exit(1)
	}

	csvFile := os.Args[1]
	docxFile := strings.TrimSuffix(csvFile, filepath.Ext(csvFile)) + ".docx"

	fmt.Printf("กำลังอ่านไฟล์ CSV: %s\n", csvFile)

	// อ่าน CSV
	chapters, err := readChapterCSV(csvFile)
	if err != nil {
		log.Fatalf("ไม่สามารถอ่านไฟล์ CSV: %v", err)
	}

	fmt.Printf("พบ %d บท\n", len(chapters))

	// Export เป็น DOCX
	fmt.Printf("กำลังสร้างไฟล์ DOCX: %s\n", docxFile)
	err = exportToDocx(chapters, docxFile)
	if err != nil {
		log.Fatalf("ไม่สามารถสร้างไฟล์ DOCX: %v", err)
	}

	fmt.Printf("✅ สำเร็จ! ไฟล์ถูกสร้างที่: %s\n", docxFile)
}

func readChapterCSV(filename string) ([]ChapterData, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	reader := csv.NewReader(file)
	
	// ปรับ configuration สำหรับ CSV ที่ซับซ้อน
	reader.LazyQuotes = true     // อนุญาตให้มี quote ที่ไม่ standard
	reader.TrimLeadingSpace = true // ตัด leading space
	reader.FieldsPerRecord = -1   // ไม่จำกัดจำนวน field ต่อ record
	
	// อ่าน CSV ทีละบรรทัด เพื่อจัดการ error ได้ดีกว่า
	var records [][]string
	for {
		record, err := reader.Read()
		if err != nil {
			if err.Error() == "EOF" {
				break
			}
			// พิมพ์ error detail และข้าม record ที่มีปัญหา
			fmt.Printf("⚠️ CSV parse error: %v - skipping line\n", err)
			continue
		}
		records = append(records, record)
	}

	if len(records) < 2 {
		return nil, fmt.Errorf("ไฟล์ CSV ต้องมีอย่างน้อย 2 แถว")
	}

	var chapters []ChapterData
	for i := 1; i < len(records); i++ {
		row := records[i]
		if len(row) < 3 {
			fmt.Printf("⚠️ Row %d has only %d columns, skipping\n", i+1, len(row))
			continue
		}

		chapters = append(chapters, ChapterData{
			ID:      strings.TrimSpace(row[0]),
			Chapter: strings.TrimSpace(row[1]),
			Body:    strings.TrimSpace(row[2]),
		})
	}

	return chapters, nil
}

func exportToDocx(chapters []ChapterData, filename string) error {
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
	if err := createStyles(zipWriter); err != nil {
		return err
	}

	// สร้าง document.xml จาก CSV data
	return createDocumentFromCSV(zipWriter, chapters)
}

func createDocumentFromCSV(zipWriter *zip.Writer, chapters []ChapterData) error {
	w, err := zipWriter.Create("word/document.xml")
	if err != nil {
		return err
	}

	doc := Document{
		Xmlns:    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		XmlnsW14: "http://schemas.microsoft.com/office/word/2010/wordml",
		Body: Body{
			Content: []interface{}{},
			SectPr: SectPr{
				PgSz:  PgSz{W: "11906", H: "16838"},
				PgMar: PgMar{Top: "1440", Right: "1440", Bottom: "1440", Left: "1440"},
			},
		},
	}

	// เพิ่มเนื้อหาแต่ละบท
	for i, chapter := range chapters {
		// Page break ก่อนบทที่ 2 เป็นต้นไป
		if i > 0 {
			pageBreak := Paragraph{
				Runs: []Run{{Break: &Break{Type: "page"}}},
			}
			doc.Body.Content = append(doc.Body.Content, pageBreak)
		}

		// หัวข้อบท - ใช้ชื่อบทจาก CSV
	// แทนบล็อกเดิมที่สร้าง title:
title := Paragraph{
  Props: &PPr{
    // กำหนดให้เป็น Heading1 เพื่อโผล่ใน Navigation Pane
    PStyle:    &PStyle{Val: "Heading1"},
    OutlineLvl:&OutlineLvl{Val: "0"},       // ระดับ 0 = หัวข้อหลัก
    Spacing:   &Spacing{Before: "480", After: "240"},
  },
  Runs: []Run{{
    Props: &RPr{
      Bold: &Bold{},       // ตัวหนาแบบเดิม
      Size: &Size{Val: "28"},
    },
    Text: &Text{
      Value: chapter.Chapter,
      Space: "preserve",
    },
  }},
}
doc.Body.Content = append(doc.Body.Content, title)


		// แปลง body content
		bodyParagraphs := convertHTMLToParagraphs(chapter.Body)
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

	if _, err = w.Write([]byte(xmlHeader)); err != nil {
		return err
	}

	_, err = io.Copy(w, &buf)
	return err
}

func convertHTMLToParagraphs(htmlContent string) []Paragraph {
	var paragraphs []Paragraph

	// ขั้นตอนที่ 1: html.UnescapeString() เพื่อแปลง HTML entities
	content := html.UnescapeString(htmlContent)

	// ขั้นตอนที่ 2: ลบ HTML comments และ special elements
	content = cleanupHTML(content)

	// ขั้นตอนที่ 3: แยก paragraphs โดยใช้ <p> tags
	pRegex := regexp.MustCompile(`<p([^>]*)>(.*?)</p>`)
	pMatches := pRegex.FindAllStringSubmatch(content, -1)

	for _, match := range pMatches {
		attributes := match[1]
		content := strings.TrimSpace(match[2])

		// จัดการ paragraph ว่างหรือมีแค่ &nbsp;
		if content == "" || isEmptyOrOnlyNbsp(content) {
			// สร้าง empty paragraph แต่รักษา attributes (เช่น class="indent-a")
			para := createEmptyParagraphWithAttributes(attributes)
			paragraphs = append(paragraphs, para)
			continue
		}

		// สร้าง paragraph ปกติ
		para := createParagraphFromHTML(content, attributes)
		paragraphs = append(paragraphs, para)
	}

	// ถ้าไม่มี <p> tags ให้สร้าง paragraph เดียว
	if len(paragraphs) == 0 {
		para := createParagraphFromHTML(content, "")
		paragraphs = append(paragraphs, para)
	}

	return paragraphs
}

// ฟังก์ชันตรวจสอบว่าเป็น empty หรือมีแค่ &nbsp;
func isEmptyOrOnlyNbsp(content string) bool {
	// แปลง HTML entities ก่อน
	unescaped := html.UnescapeString(content)
	
	// ลบ whitespace ธรรมดา
	trimmed := strings.TrimSpace(unescaped)
	
	// ตรวจสอบว่าว่างหรือมีแค่ non-breaking spaces
	if trimmed == "" {
		return true
	}
	
	// ตรวจสอบว่ามีแค่ &nbsp; โดยแปลงเป็น regular space ก่อน
	// แล้วตรวจสอบว่าเหลือแค่ whitespace หรือไม่
	withoutNbsp := strings.ReplaceAll(trimmed, "\u00A0", " ")
	finalCheck := strings.TrimSpace(withoutNbsp)
	
	return finalCheck == ""
}

// ฟังก์ชันสร้าง empty paragraph ที่รักษา attributes
func createEmptyParagraphWithAttributes(attributes string) Paragraph {
	para := Paragraph{
		Props: &PPr{
			Spacing: &Spacing{After: "120"},
		},
		Runs: []Run{
			{
				Text: &Text{Value: "", Space: "preserve"},
			},
		},
	}

	// จัดการ class="indent-a" สำหรับ empty paragraph
	if strings.Contains(attributes, `class="indent-a"`) {
		para.Props.Ind = &Ind{
			FirstLine: "720", // 0.5 inch first line indent (not left indent)
		}
	}

	// จัดการ text-align
	if regexp.MustCompile(`text-align:\s*center`).MatchString(attributes) {
		para.Props.Jc = &Jc{Val: "center"}
	}

	return para
}

func cleanupHTML(content string) string {
	// ลบ <details> (spoiler boxes)
	content = regexp.MustCompile(`<details[^>]*>.*?</details>`).ReplaceAllString(content, "")
	// ลบ <hr>
	content = regexp.MustCompile(`<hr[^>]*>`).ReplaceAllString(content, "")
	// ลบ HTML comments
	content = regexp.MustCompile(`<!--.*?-->`).ReplaceAllString(content, "")
	return content
}

func createParagraphFromHTML(content, attributes string) Paragraph {
	para := Paragraph{
		Props: &PPr{
			Spacing: &Spacing{After: "120"},
		},
		Runs: []Run{},
	}

	// Debug: แสดงการตรวจสอบ class="indent-a"

	
	// จัดการ class="indent-a"
	if hasIndentAClass(attributes) {
		if para.Props.Ind == nil {
			para.Props.Ind = &Ind{}
		}
		// ใช้ FirstLine indent แทน Left indent
		// FirstLine จะ indent เฉพาะบรรทัดแรก ส่วนบรรทัดที่ wrap จะไม่ indent
		para.Props.Ind.FirstLine = "720" // 0.5 inch first line indent
	} 

	// จัดการ text-align
	if regexp.MustCompile(`text-align:\s*center`).MatchString(attributes) {
		para.Props.Jc = &Jc{Val: "center"}
	}

	// แปลง content เป็น runs
	runs := parseContentToRuns(content)
	para.Runs = runs

	return para
}

// ฟังก์ชันตรวจสอบ indent-a แบบละเอียด
func hasIndentAClass(attributes string) bool {
	// วิธีที่ 1: ตรวจสอบแบบเดิม
	if strings.Contains(attributes, `class="indent-a"`) {
		return true
	}
	
	// วิธีที่ 2: ตรวจสอบแบบ single quotes
	if strings.Contains(attributes, `class='indent-a'`) {
		return true
	}
	
	// วิธีที่ 3: ตรวจสอบใน multiple classes
	classRegex := regexp.MustCompile(`class=["']([^"']*)["']`)
	matches := classRegex.FindStringSubmatch(attributes)
	if len(matches) > 1 {
		classes := strings.Split(matches[1], " ")
		for _, class := range classes {
			if strings.TrimSpace(class) == "indent-a" {
				return true
			}
		}
	}
	
	return false
}

func parseContentToRuns(htmlContent string) []Run {
    var runs []Run
    content := htmlContent

    // 1) แปลง <br> เป็น marker
    brRe := regexp.MustCompile(`<br\s*/?>`)
    content = brRe.ReplaceAllString(content, "___LINEBREAK___")

    // 2) จับ <span style="…"><strong>…</strong></span> (เหมือนเดิม)
    spanStrongRe := regexp.MustCompile(
        `<span[^>]*style=["']([^"']*)["'][^>]*>\s*<(strong|b)[^>]*>(.*?)</(?:strong|b)>\s*</span>`,
    )
    for spanStrongRe.MatchString(content) {
        for _, m := range spanStrongRe.FindAllStringSubmatch(content, -1) {
            props := parseColorFromStyle(m[1])
            props.Bold = &Bold{}
            runs = append(runs,
                processLineBreaksInText(processNbspAndEntities(m[3]), props)...,
            )
        }
        content = spanStrongRe.ReplaceAllString(content, "")
    }

    // 3) **จับ <span style="…"><i>…</i></span>** เพิ่มขึ้นมา
    spanItalicRe := regexp.MustCompile(
        `<span[^>]*style=["']([^"']*)["'][^>]*>\s*<(i|em)[^>]*>(.*?)</(?:i|em)>\s*</span>`,
    )
    for spanItalicRe.MatchString(content) {
        for _, m := range spanItalicRe.FindAllStringSubmatch(content, -1) {
            props := parseColorFromStyle(m[1])
            props.Italic = &Italic{}
            runs = append(runs,
                processLineBreaksInText(processNbspAndEntities(m[3]), props)...,
            )
        }
        content = spanItalicRe.ReplaceAllString(content, "")
    }

    // 4) จับ <span style="…">…</span> (ไม่มี strong/i)
    spanRe := regexp.MustCompile(`<span[^>]*style=["']([^"']*)["'][^>]*>(.*?)</span>`)
    for spanRe.MatchString(content) {
        for _, m := range spanRe.FindAllStringSubmatch(content, -1) {
            props := parseColorFromStyle(m[1])
            runs = append(runs, processLineBreaksInText(
                processNbspAndEntities(m[2]), props)...,
            )
        }
       content = spanRe.ReplaceAllString(content, "")
    }

    // 5) จับ <strong>…</strong> หรือ <b>…</b>
    strongRe := regexp.MustCompile(`<(strong|b)[^>]*>(.*?)</(?:strong|b)>`)
    for strongRe.MatchString(content) {
        for _, m := range strongRe.FindAllStringSubmatch(content, -1) {
            runs = append(runs, processLineBreaksInText(
                processNbspAndEntities(m[2]), &RPr{Bold: &Bold{}})...,
            )
        }
        content = strongRe.ReplaceAllString(content, "$2")
    }

    // 6) จับ <i>…</i> หรือ <em>…</em>
    italicRe := regexp.MustCompile(`<(i|em)[^>]*>(.*?)</(?:i|em)>`)
    for italicRe.MatchString(content) {
        for _, m := range italicRe.FindAllStringSubmatch(content, -1) {
            runs = append(runs, processLineBreaksInText(
                processNbspAndEntities(m[2]), &RPr{Italic: &Italic{}})...,
            )
        }
        content = italicRe.ReplaceAllString(content, "$2")
    }

    // 7) ลบ tag ที่เหลือ และแปลงข้อความธรรมดา
    stripped := regexp.MustCompile(`<[^>]+>`).ReplaceAllString(content, "")
    stripped = processNbspAndEntities(stripped)
    if strings.TrimSpace(stripped) != "" {
        runs = append(runs, processLineBreaksInText(stripped, nil)...)
    }

    return runs
}





// ฟังก์ชันประมวลผล &nbsp; และ HTML entities
func processNbspAndEntities(text string) string {
	// Debug: ดูว่ามี &nbsp; อยู่จริงไหม
	if strings.Contains(text, "&nbsp;") {
		fmt.Printf("🔍 Found &nbsp; in text: %q\n", text)
	}
	
	// วิธีที่ 1: แปลงทีละขั้นตอน
	// ก่อนอื่นแปลง &nbsp; เป็น special marker
	text = strings.ReplaceAll(text, "&nbsp;", "░")
	
	// แปลง HTML entities อื่นๆ
	text = html.UnescapeString(text)
	
	// แปลง marker กลับเป็น regular space (เริ่มต้นด้วย space ธรรมดาก่อน)
	text = strings.ReplaceAll(text, "░", " ")
	

	
	return text
}

// ฟังก์ชันประมวลผล line breaks ใน text
func processLineBreaksInText(text string, props *RPr) []Run {
	var runs []Run
	
	// แยก text ตาม line break marker
	parts := strings.Split(text, "___LINEBREAK___")
	
	for i, part := range parts {
		// เพิ่ม text run
		if part != "" || i == 0 { // เพิ่ม empty run สำหรับ part แรกเสมอ
			run := Run{
				Props: props,
				Text:  &Text{Value: part, Space: "preserve"},
			}
			runs = append(runs, run)
		}
		
		// เพิ่ม line break run (ยกเว้น part สุดท้าย)
		if i < len(parts)-1 {
			breakRun := Run{
				Props: props,
				Break: &Break{}, // line break (ไม่ใส่ Type สำหรับ line break ธรรมดา)
			}
			runs = append(runs, breakRun)
		}
	}
	
	return runs
}

// ฟังก์ชันประมวลผล nested tags
func processNestedTags(text string) string {
	// จัดการ <strong> nested
	text = regexp.MustCompile(`<(strong|b)[^>]*>(.*?)</(strong|b)>`).ReplaceAllString(text, "$2")
	// จัดการ <i> nested  
	text = regexp.MustCompile(`<(i|em)[^>]*>(.*?)</(i|em)>`).ReplaceAllString(text, "$2")
	// ลบ tags อื่นๆ ที่เหลือ
	text = regexp.MustCompile(`<[^>]+>`).ReplaceAllString(text, "")
	return text
}

func parseColorFromStyle(styles string) *RPr {
	rPr := &RPr{}

	// Parse #000000 format
	if matches := regexp.MustCompile(`color:\s*#([0-9a-fA-F]{6})`).FindStringSubmatch(styles); len(matches) > 1 {
		rPr.Color = &Color{Val: strings.ToUpper(matches[1])}
	}
	return rPr
}

// ฟังก์ชันสร้างไฟล์ DOCX พื้นฐาน
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
    <Application>CSV to DOCX Converter</Application>
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
    <dc:title>Document from CSV</dc:title>
    <dc:creator>CSV to DOCX Converter</dc:creator>
    <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>`

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
        <w:link w:val="Heading1Char"/>
        <w:uiPriority w:val="9"/>
        <w:qFormat/>
        <w:pPr>
            <w:keepNext/>
            <w:keepLines/>
            <w:spacing w:before="480" w:after="240"/>
            <w:outlineLvl w:val="0"/>
        </w:pPr>
        <w:rPr>
            <w:b/>
            <w:sz w:val="32"/>
            <w:szCs w:val="32"/>
        </w:rPr>
    </w:style>

    <w:style w:type="character" w:styleId="Heading1Char" w:customStyle="1">
        <w:name w:val="Heading 1 Char"/>
        <w:basedOn w:val="DefaultParagraphFont"/>
        <w:link w:val="Heading1"/>
        <w:uiPriority w:val="9"/>
        <w:rPr>
            <w:b/>
            <w:sz w:val="32"/>
            <w:szCs w:val="32"/>
        </w:rPr>
    </w:style>
</w:styles>`

	_, err = w.Write([]byte(content))
	return err
}