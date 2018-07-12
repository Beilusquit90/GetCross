// GetCross project main.go
package main

import (
	"fmt"
	"reflect"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/PuerkitoBio/goquery"
)

func _check(err error) {
	if err != nil {
		panic(err)
	}
}

func getData() { //ABDCE
	vag := strings.Split("AUDI/SEAT/SKODA/VOLKSWAGEN/SKODA (SVW )/VW (FAW)/VW (SVW)", "/")
	GM := strings.Split("CHEVROLET (SGM)/BUICK/CADILLAC/CHEVROLET/DAEWOO/OPEL/BUICK (SGM)/VAUXHALL/GM", "/")
	BMW := strings.Split("BMW/BRILLIANCE/BMW (BRILLIANCE)/MINI/ALPINA/ROLLS-ROYCE", "/")
	CP := strings.Split("CITROËN/PSA/PEUGEOT (DF-PSA)/PSA/TALBOT", "/")
	FORD := strings.Split("FORD (CHANGAN)/FORD USA/", "/")
	FIAT := strings.Split("FIAT/ALFA/LANCIA", "/")
	KIA := strings.Split("HYUNDAI/KIA", "/")
	CIT := strings.Split("CITROEN/PEUGEOT", "/")
	xlre, err := excelize.OpenFile("1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	for count := 2; len(xlre.GetCellValue("Sheet1", "B"+strconv.Itoa(count))) > 0; count++ {
		A := strings.TrimSpace(string(xlre.GetCellValue("Sheet1", "A"+strconv.Itoa(count))))
		Bb := strings.Split(strings.TrimSpace(string(xlre.GetCellValue("Sheet1", "B"+strconv.Itoa(count)))), "/")
		for _, value := range Bb {
			B := value
			D := strings.Replace(strings.Replace(strings.Replace(strings.Replace(string(xlre.GetCellValue("Sheet1", "D"+strconv.Itoa(count))), " ", "", -1), "-", "", -1), ".", "", -1), "/", "", -1)

			E := strings.TrimSpace(string(xlre.GetCellValue("Sheet1", "E"+strconv.Itoa(count))))
			dm2[B] = E
			Cb := strings.Split(strings.ToUpper(strings.TrimSpace(string(xlre.GetCellValue("Sheet1", "C"+strconv.Itoa(count))))), "/")
			for _, C := range Cb {
				fm := 0
				if C == "MERCEDES-BENZ" {
					C = "MERCEDES"
					fm = 1
				}
				if C == "SSANGYONG" {
					C = "SSANG YONG"
					fm = 1

				}
				if C == "MERCEDES-BENZ (FJDA)" {
					C = "MERCEDES"
					fm = 1
				}

				if C == "MERCEDES" {
					for _, aaa := range D {
						if aaa != 'A' {
							D = "A" + D
						}
						break
					}
				}
				for _, value := range vag {
					if C == value {
						C = "VAG"
						fm = 1
						break
					}
				}
				for _, value := range CP {
					if fm == 1 {
						break
					}
					if C == value {
						fm = 1
						C = "CITROEN/PEUGEOT"
						break
					}
				}
				for _, value := range GM {
					if fm == 1 {
						break
					}
					if C == value {
						fm = 1
						C = "GENERAL MOTORS"
						break
					}
				}
				for _, value := range BMW {
					if fm == 1 {
						break
					}
					if C == value {
						fm = 1
						C = "BMW"
						break
					}
				}
				for _, value := range FORD {
					if fm == 1 {
						break
					}
					if C == value {
						fm = 1
						C = "FORD"
						break
					}
				}
				for _, value := range FIAT {
					if fm == 1 {
						break
					}
					if C == value {
						fm = 1
						C = "FIAT/ALFA/LANCIA"
						break
					}
				}
				for _, value := range KIA {
					if fm == 1 {
						break
					}
					if C == value {
						fm = 1
						C = "HYUNDAI/KIA"
						break
					}
				}
				for _, value := range CIT {
					if fm == 1 {
						break
					}
					if C == value {
						fm = 1
						C = "CITROEN/PEUGEOT"
						break
					}
				}

				if len(dm[A]) == 0 {
					dm[A] = make(map[string]map[string]map[string]string)
				}
				if len(dm[A][B]) == 0 {
					dm[A][B] = make(map[string]map[string]string)
				}
				if len(dm[A][B][D]) == 0 {
					dm[A][B][D] = make(map[string]string)
				}
				if len(dm[A][B][D][C]) == 0 {
					dm[A][B][D][C] = E
				}
				fmt.Println(count)
			}
		}

		//fmt.Println(dm)
	}
}

func cXlsx() {
	xlre, err := excelize.OpenFile("templ/templ.xlsx")
	count := 2
	for _, Ax := range reflect.ValueOf(dm).MapKeys() {
		A := Ax.Interface().(string)
		for _, Bx := range reflect.ValueOf(dm[A]).MapKeys() {
			B := Bx.Interface().(string)
			for _, Dx := range reflect.ValueOf(dm[A][B]).MapKeys() {
				D := Dx.Interface().(string)
				for _, Cx := range reflect.ValueOf(dm[A][B][D]).MapKeys() {
					C := Cx.Interface().(string)
					nname := ""
					if len(dm2[B]) == 0 {
						nname = string(dm[A][B][D][C])
					} else {
						nname = dm2[B]
					}
					if len(string(D)) < 1 {
						continue
					}
					xlre.SetCellValue("Sheet1", "A"+strconv.Itoa(count), string(A))
					xlre.SetCellValue("Sheet1", "B"+strconv.Itoa(count), string(B))
					xlre.SetCellValue("Sheet1", "D"+strconv.Itoa(count), string(D))
					xlre.SetCellValue("Sheet1", "C"+strconv.Itoa(count), string(C))
					xlre.SetCellValue("Sheet1", "E"+strconv.Itoa(count), string(nname))
					count++
				}
			}
		}
	}
	err = xlre.SaveAs("NewCross.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}

func parseUrl(url string, firm string) {
	vag := strings.Split("AUDI/SEAT/SKODA/VOLKSWAGEN/SKODA (SVW )/VW (FAW)/VW (SVW)", "/")
	GM := strings.Split("CHEVROLET (SGM)/BUICK/CADILLAC/CHEVROLET/DAEWOO/OPEL/BUICK (SGM)/VAUXHALL/GM", "/")
	BMW := strings.Split("BMW/BRILLIANCE/BMW (BRILLIANCE)/MINI/ALPINA/ROLLS-ROYCE", "/")
	CP := strings.Split("CITROËN/PSA/PEUGEOT (DF-PSA)/PSA/TALBOT", "/")
	FORD := strings.Split("FORD (CHANGAN)/FORD USA/", "/")
	FIAT := strings.Split("FIAT/ALFA/LANCIA", "/")
	KIA := strings.Split("HYUNDAI/KIA", "/")
	CIT := strings.Split("CITROEN/PEUGEOT", "/")
	fmt.Println("request: " + url)
	doc, err := goquery.NewDocument("http://webcat.borgandbeck.com/PartDetails/" + url + "/#partInfo" + url)
	_check(err)
	flag := 0
	fl := 0
	td := []string{}
	doc.Find(".col-xs-4").Each(func(i int, s *goquery.Selection) {
		flag++
		if flag > 3 {
			fl++
			link := s.Text()
			if fl < 3 {
				td = append(td, link)
			} else {
				A := firm
				B := url
				D := strings.Replace(strings.Replace(strings.Replace(strings.Replace(td[1], " ", "", -1), "-", "", -1), ".", "", -1), "/", "", -1)
				E := ""
				co := 1
				doc.Find("dd").Each(func(q int, z *goquery.Selection) {
					if co == 1 {
						E = z.Text()
						co++
					}
				})
				//fmt.Println(E)
				C := strings.TrimSpace(strings.ToUpper(td[0]))
				fm := 0

				if C == "MERCEDES-BENZ" {
					C = "MERCEDES"
					fm = 1
				}
				if C == "SSANGYONG" {
					C = "SSANG YONG"
					fm = 1

				}
				if C == "MERCEDES-BENZ (FJDA)" {
					C = "MERCEDES"
					fm = 1
				}
				if C == "MERCEDES" {
					for _, aaa := range D {
						if aaa != 'A' {
							D = "A" + D
						}
						break
					}
				}
				for _, value := range vag {
					if C == value {
						C = "VAG"
						fm = 1
						break
					}
				}
				for _, value := range CP {
					if fm == 1 {
						break
					}
					if C == value {
						fm = 1
						C = "CITROEN/PEUGEOT"
						break
					}
				}
				for _, value := range GM {
					if fm == 1 {
						break
					}
					if C == value {
						fm = 1
						C = "GENERAL MOTORS"
						break
					}
				}
				for _, value := range BMW {
					if fm == 1 {
						break
					}
					if C == value {
						fm = 1
						C = "BMW"
						break
					}
				}
				for _, value := range FORD {
					if fm == 1 {
						break
					}
					if C == value {
						fm = 1
						C = "FORD"
						break
					}
				}
				for _, value := range FIAT {
					if fm == 1 {
						break
					}
					if C == value {
						fm = 1
						C = "FIAT/ALFA/LANCIA"
						break
					}
				}
				for _, value := range KIA {
					if fm == 1 {
						break
					}
					if C == value {
						fm = 1
						C = "HYUNDAI/KIA"
						break
					}
				}
				for _, value := range CIT {
					if fm == 1 {
						break
					}
					if C == value {
						fm = 1
						C = "CITROEN/PEUGEOT"
						break
					}
				}

				if len(dm[A]) == 0 {
					dm[A] = make(map[string]map[string]map[string]string)
				}
				if len(dm[A][B]) == 0 {
					dm[A][B] = make(map[string]map[string]string)
				}
				if len(dm[A][B][D]) == 0 {
					dm[A][B][D] = make(map[string]string)
				}
				if len(dm[A][B][D][C]) == 0 {
					dm[A][B][D][C] = E
				}
				td = []string{}
				fl = 0
			}
		}
	})
}

var dm map[string]map[string]map[string]map[string]string
var dm2 map[string]string

func main() {
	dm2 = make(map[string]string)
	dm = make(map[string]map[string]map[string]map[string]string)
	getData()
	var wg sync.WaitGroup
	for _, f2 := range reflect.ValueOf(dm).MapKeys() {
		firm := f2.Interface().(string)
		for _, url2 := range reflect.ValueOf(dm[firm]).MapKeys() {
			url := url2.Interface().(string)
			wg.Add(1)
			time.Sleep(30 * time.Millisecond)
			go func(url string, firm string) {
				defer wg.Done()
				parseUrl(url, firm)
			}(url, firm)
		}
	}
	wg.Wait()
	fmt.Println("НАЧАЛАСЬ ГЕНЕРАЦИЯ ФАЙЛА! просто подожди немного пока программа закроется и появится файл.")
	cXlsx()
}
