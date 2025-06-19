package main

import (
	"archive/zip"
	"bytes"
	"crypto/md5"
	"encoding/csv"
	"encoding/hex"
	"encoding/xml"
	"fmt"
	"html"
	"io"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

// โครงสร้างข้อมูลสำหรับ CSV
type ChapterData struct {
	ID      string
	Chapter string
	Body    string
}

// DOCX XML Structures
type Document struct {
	XMLName  xml.Name `xml:"w:document"`
	Xmlns    string   `xml:"xmlns:w,attr"`
	XmlnsR   string   `xml:"xmlns:r,attr"`
	XmlnsWP  string   `xml:"xmlns:wp,attr"`
	XmlnsA   string   `xml:"xmlns:a,attr"`
	XmlnsPic string   `xml:"xmlns:pic,attr"`
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
	Drawing *Drawing `xml:"w:drawing,omitempty"`
}

type Drawing struct {
	XMLName xml.Name `xml:"w:drawing"`
	Inline  *Inline  `xml:"wp:inline"`
}

type Inline struct {
	XMLName    xml.Name `xml:"wp:inline"`
	DistT      string   `xml:"distT,attr"`
	DistB      string   `xml:"distB,attr"`
	DistL      string   `xml:"distL,attr"`
	DistR      string   `xml:"distR,attr"`
	Extent     Extent   `xml:"wp:extent"`
	EffectExt  EffectExt `xml:"wp:effectExtent"`
	DocPr      DocPr    `xml:"wp:docPr"`
	CNvGraphicFramePr CNvGraphicFramePr `xml:"wp:cNvGraphicFramePr"`
	Graphic    Graphic  `xml:"a:graphic"`
}

type Extent struct {
	XMLName xml.Name `xml:"wp:extent"`
	Cx      string   `xml:"cx,attr"`
	Cy      string   `xml:"cy,attr"`
}

type EffectExt struct {
	XMLName xml.Name `xml:"wp:effectExtent"`
	L       string   `xml:"l,attr"`
	T       string   `xml:"t,attr"`
	R       string   `xml:"r,attr"`
	B       string   `xml:"b,attr"`
}

type DocPr struct {
	XMLName xml.Name `xml:"wp:docPr"`
	Id      string   `xml:"id,attr"`
	Name    string   `xml:"name,attr"`
}

type CNvGraphicFramePr struct {
	XMLName xml.Name `xml:"wp:cNvGraphicFramePr"`
	GraphicFrameLocks GraphicFrameLocks `xml:"a:graphicFrameLocks"`
}

type GraphicFrameLocks struct {
	XMLName         xml.Name `xml:"a:graphicFrameLocks"`
	NoChangeAspect  string   `xml:"noChangeAspect,attr"`
}

type Graphic struct {
	XMLName    xml.Name    `xml:"a:graphic"`
	GraphicData GraphicData `xml:"a:graphicData"`
}

type GraphicData struct {
	XMLName xml.Name `xml:"a:graphicData"`
	Uri     string   `xml:"uri,attr"`
	Pic     Pic      `xml:"pic:pic"`
}

type Pic struct {
	XMLName   xml.Name   `xml:"pic:pic"`
	NvPicPr   NvPicPr    `xml:"pic:nvPicPr"`
	BlipFill  BlipFill   `xml:"pic:blipFill"`
	SpPr      SpPr       `xml:"pic:spPr"`
}

type NvPicPr struct {
	XMLName xml.Name `xml:"pic:nvPicPr"`
	CNvPr   CNvPr    `xml:"pic:cNvPr"`
	CNvPicPr CNvPicPr `xml:"pic:cNvPicPr"`
}

type CNvPr struct {
	XMLName xml.Name `xml:"pic:cNvPr"`
	Id      string   `xml:"id,attr"`
	Name    string   `xml:"name,attr"`
}

type CNvPicPr struct {
	XMLName xml.Name `xml:"pic:cNvPicPr"`
}

type BlipFill struct {
	XMLName xml.Name `xml:"pic:blipFill"`
	Blip    Blip     `xml:"a:blip"`
	Stretch Stretch  `xml:"a:stretch"`
}

type Blip struct {
	XMLName xml.Name `xml:"a:blip"`
	Embed   string   `xml:"r:embed,attr"`
}

type Stretch struct {
	XMLName  xml.Name  `xml:"a:stretch"`
	FillRect FillRect  `xml:"a:fillRect"`
}

type FillRect struct {
	XMLName xml.Name `xml:"a:fillRect"`
}

type SpPr struct {
	XMLName xml.Name `xml:"pic:spPr"`
	Xfrm    Xfrm     `xml:"a:xfrm"`
	PrstGeom PrstGeom `xml:"a:prstGeom"`
}

type Xfrm struct {
	XMLName xml.Name `xml:"a:xfrm"`
	Off     Off      `xml:"a:off"`
	Ext     Ext      `xml:"a:ext"`
}

type Off struct {
	XMLName xml.Name `xml:"a:off"`
	X       string   `xml:"x,attr"`
	Y       string   `xml:"y,attr"`
}

type Ext struct {
	XMLName xml.Name `xml:"a:ext"`
	Cx      string   `xml:"cx,attr"`
	Cy      string   `xml:"cy,attr"`
}

type PrstGeom struct {
	XMLName xml.Name `xml:"a:prstGeom"`
	Prst    string   `xml:"prst,attr"`
	AvLst   AvLst    `xml:"a:avLst"`
}

type AvLst struct {
	XMLName xml.Name `xml:"a:avLst"`
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

// โครงสร้างสำหรับ relationships
type Relationship struct {
	Id     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

type Relationships struct {
	XMLName xml.Name `xml:"Relationships"`
	Xmlns   string   `xml:"xmlns,attr"`
	Items   []Relationship `xml:"Relationship"`
}

// โครงสร้างสำหรับจัดการรูปภาพ
type ImageInfo struct {
	URL      string
	Data     []byte
	Filename string
	RelId    string
	Width    int
	Height   int
}

// ตัวแปรสำหรับเก็บรูปภาพ
var (
	imageCounter = 1
	images       []ImageInfo
	relCounter   = 2 // เริ่มจาก 2 เพราะ rId1 ใช้กับ styles.xml
)

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

	// Reset ตัวแปรสำหรับรูปภาพ
	imageCounter = 1
	images = []ImageInfo{}
	relCounter = 2

	// Export เป็น DOCX
	fmt.Printf("กำลังสร้างไฟล์ DOCX: %s\n", docxFile)
	err = exportToDocx(chapters, docxFile)
	if err != nil {
		log.Fatalf("ไม่สามารถสร้างไฟล์ DOCX: %v", err)
	}

	fmt.Printf("✅ สำเร็จ! ไฟล์ถูกสร้างที่: %s\n", docxFile)
	if len(images) > 0 {
		fmt.Printf("📷 โหลดรูปภาพ %d รูป\n", len(images))
	}
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
	if err := createDocumentFromCSV(zipWriter, chapters); err != nil {
		return err
	}

	// สร้าง document.xml.rels สำหรับรูปภาพ
	if err := createDocumentRels(zipWriter); err != nil {
		return err
	}

	// เพิ่มรูปภาพลงใน ZIP
	if err := addImagesToZip(zipWriter); err != nil {
		return err
	}

	return nil
}

func createDocumentFromCSV(zipWriter *zip.Writer, chapters []ChapterData) error {
	w, err := zipWriter.Create("word/document.xml")
	if err != nil {
		return err
	}

	doc := Document{
		Xmlns:    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		XmlnsR:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		XmlnsWP:  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
		XmlnsA:   "http://schemas.openxmlformats.org/drawingml/2006/main",
		XmlnsPic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
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
		title := Paragraph{
			Props: &PPr{
				// กำหนดให้เป็น Heading1 เพื่อโผล่ใน Navigation Pane
				PStyle:     &PStyle{Val: "Heading1"},
				OutlineLvl: &OutlineLvl{Val: "0"},       // ระดับ 0 = หัวข้อหลัก
				Spacing:    &Spacing{Before: "480", After: "240"},
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

func convertHTMLToParagraphs(htmlContent string) []interface{} {
	var paragraphs []interface{}

	// ขั้นตอนที่ 1: html.UnescapeString() เพื่อแปลง HTML entities
	content := html.UnescapeString(htmlContent)

	// ขั้นตอนที่ 2: ลบ HTML comments และ special elements (ยกเว้น figure)
	content = cleanupHTML(content)

	// ขั้นตอนที่ 3: แยก content เป็น segments โดยคำนึงถึงตำแหน่งของ figure
	segments := parseContentWithFigures(content)

	// ขั้นตอนที่ 4: แปลงแต่ละ segment
	for _, segment := range segments {
		if segment.Type == "figure" {
			// สร้าง image paragraph
			imageParagraph := createImageParagraph(segment.ImageInfo)
			paragraphs = append(paragraphs, imageParagraph)
		} else if segment.Type == "text" {
			// แยก paragraphs โดยใช้ <p> tags
			pRegex := regexp.MustCompile(`<p([^>]*)>(.*?)</p>`)
			pMatches := pRegex.FindAllStringSubmatch(segment.Content, -1)

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
			if len(pMatches) == 0 && strings.TrimSpace(segment.Content) != "" {
				para := createParagraphFromHTML(segment.Content, "")
				paragraphs = append(paragraphs, para)
			}
		}
	}

	return paragraphs
}

// โครงสร้างสำหรับ content segment
type ContentSegment struct {
	Type      string    // "text" หรือ "figure"
	Content   string    // สำหรับ text
	ImageInfo ImageInfo // สำหรับ figure
}

// ฟังก์ชันแยก content เป็น segments ตามลำดับ
func parseContentWithFigures(content string) []ContentSegment {
	var segments []ContentSegment
	
	// จับ <figure> tags ที่มีรูปภาพ
	figureRegex := regexp.MustCompile(`<figure[^>]*class="[^"]*image[^"]*"[^>]*style="[^"]*width:\s*(\d+)%[^"]*"[^>]*>\s*<img[^>]*src="([^"]+)"[^>]*>\s*</figure>`)
	
	// หา index ของ figure tags ทั้งหมด
	figureMatches := figureRegex.FindAllStringSubmatchIndex(content, -1)
	
	lastIndex := 0
	
	for _, match := range figureMatches {
		start := match[0]
		end := match[1]
		widthPercent := content[match[2]:match[3]]
		imageURL := content[match[4]:match[5]]
		
		// เพิ่ม text segment ก่อน figure (ถ้ามี)
		if start > lastIndex {
			textContent := content[lastIndex:start]
			if strings.TrimSpace(textContent) != "" {
				segments = append(segments, ContentSegment{
					Type:    "text",
					Content: textContent,
				})
			}
		}
		
		// โหลดรูปภาพ
		imageInfo, err := downloadImage(imageURL, widthPercent)
		if err != nil {
			fmt.Printf("❌ Error downloading image %s: %v\n", imageURL, err)
			// ถ้าโหลดรูปไม่ได้ ข้าม figure นี้
			lastIndex = end
			continue
		}
		
		// เพิ่ม figure segment
		segments = append(segments, ContentSegment{
			Type:      "figure", 
			ImageInfo: imageInfo,
		})
		
		fmt.Printf("📷 Added image: %s (width: %s%%)\n", imageURL, widthPercent)
		
		lastIndex = end
	}
	
	// เพิ่ม text segment ท้ายสุด (ถ้ามี)
	if lastIndex < len(content) {
		textContent := content[lastIndex:]
		if strings.TrimSpace(textContent) != "" {
			segments = append(segments, ContentSegment{
				Type:    "text",
				Content: textContent,
			})
		}
	}
	
	// ถ้าไม่มี figure เลย ให้ส่งคืน text segment เดียว
	if len(segments) == 0 && strings.TrimSpace(content) != "" {
		segments = append(segments, ContentSegment{
			Type:    "text",
			Content: content,
		})
	}
	
	return segments
}

func downloadImage(url, widthPercent string) (ImageInfo, error) {
	// ทำความสะอาด URL
	url = html.UnescapeString(url)
	
	fmt.Printf("🔄 Downloading image: %s\n", url)
	
	// ดาวน์โหลดรูปภาพ
	resp, err := http.Get(url)
	if err != nil {
		return ImageInfo{}, err
	}
	defer resp.Body.Close()
	
	// ตรวจสอบ status code
	if resp.StatusCode != 200 {
		return ImageInfo{}, fmt.Errorf("HTTP %d: %s", resp.StatusCode, resp.Status)
	}
	
	// อ่านข้อมูลรูปภาพ
	imageData, err := io.ReadAll(resp.Body)
	if err != nil {
		return ImageInfo{}, err
	}
	
	// สร้างชื่อไฟล์จาก URL hash
	hasher := md5.New()
	hasher.Write([]byte(url))
	hash := hex.EncodeToString(hasher.Sum(nil))
	
	// กำหนดนามสกุลไฟล์ตาม Content-Type
	contentType := resp.Header.Get("Content-Type")
	ext := ".jpg" // default
	if strings.Contains(contentType, "png") {
		ext = ".png"
	} else if strings.Contains(contentType, "gif") {
		ext = ".gif"
	} else if strings.Contains(contentType, "webp") {
		ext = ".jpg" // แปลง webp เป็น jpg
	}
	
	filename := fmt.Sprintf("image%d_%s%s", imageCounter, hash[:8], ext)
	
	// คำนวณขนาดรูป - ใช้ขนาดที่เหมาะสมกับการแสดงผลใน Word
	width, height := 500, 375 // ขนาดเริ่มต้นที่ใหญ่ขึ้น (4:3 ratio)
	
	// ปรับขนาดตาม widthPercent
	if wp, err := strconv.Atoi(widthPercent); err == nil {
		scale := float64(wp) / 100.0
		// กำหนดขนาดสูงสุดที่ 600px สำหรับความกว้าง
		maxWidth := 600
		width = int(float64(maxWidth) * scale)
		height = width * 3 / 4 // รักษา aspect ratio 4:3
	}
	
	// สร้าง relationship ID
	relId := fmt.Sprintf("rId%d", relCounter)
	relCounter++
	
	imageInfo := ImageInfo{
		URL:      url,
		Data:     imageData,
		Filename: filename,
		RelId:    relId,
		Width:    width,
		Height:   height,
	}
	
	images = append(images, imageInfo)
	imageCounter++
	
	return imageInfo, nil
}

func createImageParagraph(imageInfo ImageInfo) Paragraph {
	// คำนวณขนาดใน EMU (English Metric Units)
	// 1 inch = 914400 EMU, 1 pixel = 9525 EMU (at 96 DPI)
	widthEMU := imageInfo.Width * 9525
	heightEMU := imageInfo.Height * 9525
	
	drawing := &Drawing{
		Inline: &Inline{
			DistT: "0",
			DistB: "0", 
			DistL: "0",
			DistR: "0",
			Extent: Extent{
				Cx: strconv.Itoa(widthEMU),
				Cy: strconv.Itoa(heightEMU),
			},
			EffectExt: EffectExt{
				L: "0",
				T: "0", 
				R: "0",
				B: "0",
			},
			DocPr: DocPr{
				Id:   strconv.Itoa(imageCounter),
				Name: imageInfo.Filename,
			},
			CNvGraphicFramePr: CNvGraphicFramePr{
				GraphicFrameLocks: GraphicFrameLocks{
					NoChangeAspect: "1",
				},
			},
			Graphic: Graphic{
				GraphicData: GraphicData{
					Uri: "http://schemas.openxmlformats.org/drawingml/2006/picture",
					Pic: Pic{
						NvPicPr: NvPicPr{
							CNvPr: CNvPr{
								Id:   "0",
								Name: imageInfo.Filename,
							},
							CNvPicPr: CNvPicPr{},
						},
						BlipFill: BlipFill{
							Blip: Blip{
								Embed: imageInfo.RelId,
							},
							Stretch: Stretch{
								FillRect: FillRect{},
							},
						},
						SpPr: SpPr{
							Xfrm: Xfrm{
								Off: Off{X: "0", Y: "0"},
								Ext: Ext{
									Cx: strconv.Itoa(widthEMU),
									Cy: strconv.Itoa(heightEMU),
								},
							},
							PrstGeom: PrstGeom{
								Prst:  "rect",
								AvLst: AvLst{},
							},
						},
					},
				},
			},
		},
	}
	
	return Paragraph{
		Props: &PPr{
			Jc:      &Jc{Val: "center"}, // จัดกลาง
			Spacing: &Spacing{After: "240"},
		},
		Runs: []Run{{
			Drawing: drawing,
		}},
	}
}

func createDocumentRels(zipWriter *zip.Writer) error {
	w, err := zipWriter.Create("word/_rels/document.xml.rels")
	if err != nil {
		return err
	}

	relationships := Relationships{
		Xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
		Items: []Relationship{},
	}

	// เพิ่ม relationship สำหรับ styles.xml
	relationships.Items = append(relationships.Items, Relationship{
		Id:     "rId1",
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
		Target: "styles.xml",
	})

	// เพิ่ม relationships สำหรับรูปภาพ
	for _, img := range images {
		relationships.Items = append(relationships.Items, Relationship{
			Id:     img.RelId,
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
			Target: "media/" + img.Filename,
		})
	}

	// เขียน XML
	xmlHeader := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"
	
	var buf bytes.Buffer
	encoder := xml.NewEncoder(&buf)
	encoder.Indent("", "  ")

	if err := encoder.Encode(relationships); err != nil {
		return err
	}

	if _, err = w.Write([]byte(xmlHeader)); err != nil {
		return err
	}

	_, err = io.Copy(w, &buf)
	return err
}

func addImagesToZip(zipWriter *zip.Writer) error {
	for _, img := range images {
		w, err := zipWriter.Create("word/media/" + img.Filename)
		if err != nil {
			return err
		}
		
		if _, err = w.Write(img.Data); err != nil {
			return err
		}
	}
	return nil
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
    <Default Extension="png" ContentType="image/png"/>
    <Default Extension="jpg" ContentType="image/jpeg"/>
    <Default Extension="jpeg" ContentType="image/jpeg"/>
    <Default Extension="gif" ContentType="image/gif"/>
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
        <w:rFonts
          w:ascii="Times New Roman"
          w:hAnsi="Times New Roman"
          w:cs="Times New Roman"
          w:eastAsia="TH SarabunPSK"/>
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