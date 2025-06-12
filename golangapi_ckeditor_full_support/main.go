package main

import (
	"archive/zip"
	"log"
	"os"
	"strings"

	"golang.org/x/net/html"
)

func main() {
	files, err := os.ReadDir("novel")
	if err != nil {
		log.Fatal("❌ ไม่พบโฟลเดอร์ novel/:", err)
	}

	var allParagraphs []string
	for idx, file := range files {
		if file.IsDir() {
			continue
		}
		data, _ := os.ReadFile("novel/" + file.Name())
		htmlContent := string(data)
		htmlNode, _ := html.Parse(strings.NewReader(htmlContent))
		allParagraphs = append(allParagraphs, parseDOM(htmlNode))

		if idx < len(files)-1 {
			allParagraphs = append(allParagraphs, `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`)
		}
	}

	docXML := wrapDocument(strings.Join(allParagraphs, "\n"))

	f, _ := os.Create("output.docx")
	zipWriter := zip.NewWriter(f)
	writeZip(zipWriter, "[Content_Types].xml", contentTypesXML)
	writeZip(zipWriter, "_rels/.rels", relsXML)
	writeZip(zipWriter, "word/_rels/document.xml.rels", emptyRels)
	writeZip(zipWriter, "word/document.xml", docXML)
	zipWriter.Close()
	log.Println("✅ สร้างไฟล์ output.docx สำเร็จ")
}

type StyledRun struct {
	Text      string
	Bold      bool
	Italic    bool
	Underline bool
}

func parseDOM(n *html.Node) string {
	var result []string

	var walk func(*html.Node)
	walk = func(n *html.Node) {
		if n.Type == html.ElementNode {
			switch n.Data {
			case "p", "h1", "h2":
				className := getClassAttr(n)
				runs := extractStyledRuns(n)
				result = append(result, wrapParagraph(runs, n.Data, className))

			case "ul":
				for c := n.FirstChild; c != nil; c = c.NextSibling {
					if c.Type == html.ElementNode && c.Data == "li" {
						runs := extractStyledRuns(c)
						result = append(result, wrapParagraph(runs, "li", ""))
					}
				}
			}
		}
		for c := n.FirstChild; c != nil; c = c.NextSibling {
			walk(c)
		}
	}
	walk(n)

	return strings.Join(result, "\n")
}

func extractStyledRuns(n *html.Node) []StyledRun {
	var runs []StyledRun
	var walk func(*html.Node, StyledRun)

	walk = func(n *html.Node, current StyledRun) {
		if n.Type == html.ElementNode {
			switch n.Data {
			case "b", "strong":
				current.Bold = true
			case "i", "em":
				current.Italic = true
			case "u":
				current.Underline = true
			}
		} else if n.Type == html.TextNode {
			current.Text = n.Data
			runs = append(runs, current)
		}
		for c := n.FirstChild; c != nil; c = c.NextSibling {
			walk(c, current)
		}
	}
	walk(n, StyledRun{})
	return runs
}

func wrapParagraph(runs []StyledRun, tag, className string) string {
	pPr := ""
	if strings.Contains(className, "indent") {
		pPr += `<w:ind w:left="720"/>`
	}
	if tag == "h1" || tag == "h2" {
		pPr += `<w:spacing w:after="200"/>`
	}
	result := "<w:p><w:pPr>" + pPr + "</w:pPr>"
	for _, run := range runs {
		rPr := ""
		if tag == "h1" || tag == "h2" || run.Bold {
			rPr += "<w:b/>"
		}
		if run.Italic {
			rPr += "<w:i/>"
		}
		if run.Underline {
			rPr += `<w:u w:val="single"/>`
		}
		rPrXML := ""
		if rPr != "" {
			rPrXML = "<w:rPr>" + rPr + "</w:rPr>"
		}
		result += `<w:r>` + rPrXML + `<w:t xml:space="preserve">` + run.Text + `</w:t></w:r>`
	}
	result += "</w:p>"
	return result
}



func extractText(n *html.Node) string {
	var b strings.Builder
	var f func(*html.Node)
	f = func(n *html.Node) {
		if n.Type == html.TextNode {
			b.WriteString(n.Data)
		}
		for c := n.FirstChild; c != nil; c = c.NextSibling {
			f(c)
		}
	}
	f(n)
	return b.String()
}

func getStyle(n *html.Node) string {
	for _, attr := range n.Attr {
		if attr.Key == "style" {
			return attr.Val
		}
	}
	return ""
}

func getClassAttr(n *html.Node) string {
	for _, attr := range n.Attr {
		if attr.Key == "class" {
			return attr.Val
		}
	}
	return ""
}

func wrapStyledText(text, style string) string {
	rPr := ""
	if strings.Contains(style, "text-align:center") {
		return `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>` + text + `</w:t></w:r></w:p>`
	}
	if strings.Contains(style, "color") || strings.Contains(style, "background") {
		rPr += "<w:rPr><w:b/></w:rPr>" // mock bold style
	}
	return `<w:p><w:r>` + rPr + `<w:t xml:space="preserve">` + text + `</w:t></w:r></w:p>`
}

func wrapDocument(body string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ` + body + `
  </w:body>
</w:document>`
}

func writeZip(w *zip.Writer, name, content string) {
	f, _ := w.Create(name)
	f.Write([]byte(content))
}

var contentTypesXML = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`

var relsXML = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`

var emptyRels = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`
