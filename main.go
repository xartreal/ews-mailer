package main

import (
	"encoding/base64"
	"io/ioutil"
	"log"
)

func main() {
	//init mailer
	from = "alice@example.com"
	endpoint = "https://mail.example.com/ews/Exchange.asmx"
	username = "alice.e"
	userpass = "12345"

	// send text only letter, receipt enabled
	SendTextOnly("bob@example.com", "Bob Smith", "Test letter title", "Hello!\nThis is test letter body", "enabled")
	//init attachment pool
	filenames = append(filenames, "sample.docx", "test.xlsx")
	filecount = len(filenames)
	for _, fileitem := range filenames {
		fbin, err := ioutil.ReadFile(fileitem)
		if err != nil {
			log.Fatalf("File %s not found\n", fileitem)
		}
		// convert file to base64
		FileContent[fileitem] = base64.StdEncoding.EncodeToString(fbin)
	}
	//add letter to exchange, receipt disabled
	mmid, mmkey := SendLetterStep("bob@example.com", "Bob Smith", "Test w/attach", "Hello!\nThis is test letter with attachments", "disabled")
	if (len(mmid) < 2) || (len(mmkey) < 2) { //if no keys returned, send failed
		log.Fatalf("Add letter error\n")
	}
	//add attachments
	mmkey = SendAttachStep(mmid, mmkey, filenames)
	if len(mmkey) < 2 { //if no keys returned, send failed
		log.Fatalf("Add attachment error\n")
	}
	// send letter (finalize)
	SendLetterFinal(mmid, mmkey)
	if s3err {
		log.Printf("ERROR: %s\n", lasterr)
	}

}
