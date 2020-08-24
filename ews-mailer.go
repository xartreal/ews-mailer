// mailer
package main

import (
	"bytes"
	//	"encoding/base64"
	"crypto/tls"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"regexp"
	"strings"
	"time"

	"github.com/vadimi/go-http-ntlm"
)

var TplCheckAccess = `
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
        xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
        xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
        xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2010_SP1" />
  </soap:Header>
  <soap:Body>
    <m:GetFolder>
      <m:FolderShape>
        <t:BaseShape>AllProperties</t:BaseShape>
      </m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="msgfolderroot" />
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>
</soap:Envelope>
`
var TplSendRC = `
          <t:IsDeliveryReceiptRequested>true</t:IsDeliveryReceiptRequested>
`

var TplSendText = `
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <CreateItem MessageDisposition="%mailmode%" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
      <SavedItemFolderId>
        <t:DistinguishedFolderId Id="sentitems" >
          <t:Mailbox>
           <t:EmailAddress>%mailfrom%</t:EmailAddress>
          </t:Mailbox>
        </t:DistinguishedFolderId>
      </SavedItemFolderId>
      <Items>
        <t:Message>
          <t:ItemClass>IPM.Note</t:ItemClass>
          <t:Subject>%mailsubj%</t:Subject>
          <t:Body BodyType="Text">%mailtext%</t:Body>
<t:Sender>
 <t:Mailbox>
  <t:EmailAddress>%mailfrom%</t:EmailAddress>
 </t:Mailbox>
</t:Sender>     

          <t:ToRecipients>
            <t:Mailbox>
              <t:Name>%toname%</t:Name>
              <t:EmailAddress>%mailto%</t:EmailAddress>
            </t:Mailbox>
          </t:ToRecipients>%rc%
          <t:IsRead>false</t:IsRead>
        </t:Message>
      </Items>
    </CreateItem>
  </soap:Body>
</soap:Envelope>
`
var TplSendAttach = `
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <CreateAttachment xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
                      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <ParentItemId Id="%mailid%" ChangeKey="%mailkey%"/>
      <Attachments>
    %mailfiles%
      </Attachments>
    </CreateAttachment>
  </soap:Body>
</soap:Envelope>
`
var TplAttachItem = `
        <t:FileAttachment>
          <t:Name>%filename%</t:Name>
          <t:Content>%filecontent%</t:Content>
        </t:FileAttachment>
`
var TplSendFinal = `
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <SendItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
              SaveItemToFolder="true">
      <ItemIds>
        <t:ItemId Id="%mailid%" ChangeKey="%mailkey%" />
      </ItemIds>
    </SendItem>
  </soap:Body>
</soap:Envelope>
`

func senditem(xmlin string, fname string) (string, error) {
	var resp *http.Response
	var err error

	ioutil.WriteFile(fname+".log", []byte(xmlin), 0755)
	//	fmt.Printf("%s\n", xmlin)
	req, err := http.NewRequest("POST", endpoint, bytes.NewReader([]byte(xmlin)))
	if err != nil {
		log.Printf("Z0 error\n")
		return "", err
	}
	//	fmt.Printf("%v\n", req)
	req.Header.Set("Content-Type", "text/xml")
	client := http.Client{
		Transport: &httpntlm.NtlmTransport{
			Domain:          "",
			User:            username,
			Password:        userpass,
			TLSClientConfig: &tls.Config{InsecureSkipVerify: true},
		},
	}
	client.CheckRedirect = func(req *http.Request, via []*http.Request) error { return http.ErrUseLastResponse }
	//	resp, err := "", nil
	att := 0
	for {
		resp, err = client.Do(req)
		if err != nil {
			att++
			log.Printf("Attempt %d; Z1 error: %q\n", att, err.Error())
			if att > 4 {
				log.Printf("FATAL: Skipping via fatal response\n")
				return "", err
			}
			time.Sleep(1 * time.Minute)
		} else {
			break
		}
	}
	defer resp.Body.Close()
	//fmt.Printf("Status=%v\n", resp.StatusCode)
	if resp.StatusCode != 200 {
		return "", fmt.Errorf("Status=%v", resp.StatusCode)
	}
	bodyBytes, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		log.Printf("Z2 error\n")
		return "", err
	}
	bodyString := string(bodyBytes)

	return bodyString, nil
}

func CheckCR() error {
	_, err := senditem(TplCheckAccess, "cr")
	return err
}

func SendTextOnly(to string, nameto string, subj string, text string, rc string) {
	mailxml := TplSendText
	mailxml = strings.Replace(mailxml, "%mailmode%", "SendAndSaveCopy", -1)
	mailxml = strings.Replace(mailxml, "%mailto%", to, -1)
	mailxml = strings.Replace(mailxml, "%toname%", nameto, -1)
	mailxml = strings.Replace(mailxml, "%mailfrom%", from, -1)
	mailxml = strings.Replace(mailxml, "%mailsubj%", subj, -1)
	mailxml = strings.Replace(mailxml, "%mailtext%", text, -1)
	if rc == "1" {
		mailxml = strings.Replace(mailxml, "%rc%", TplSendRC, -1)
	} else {
		mailxml = strings.Replace(mailxml, "%rc%", "", -1)
	}

	resp, err := senditem(mailxml, "r0")
	if err != nil {
		log.Printf("Error %q\n", err)
	}
	ioutil.WriteFile("s0.log", []byte(resp), 0755)
}

// return item-id, item-key
func SendLetterStep(to string, nameto string, subj string, text string, rc string) (string, string) {
	mailxml := TplSendText
	mailxml = strings.Replace(mailxml, "%mailmode%", "SaveOnly", -1)
	mailxml = strings.Replace(mailxml, "%mailto%", to, -1)
	mailxml = strings.Replace(mailxml, "%toname%", nameto, -1)
	mailxml = strings.Replace(mailxml, "%mailfrom%", from, -1)
	mailxml = strings.Replace(mailxml, "%mailsubj%", subj, -1)
	mailxml = strings.Replace(mailxml, "%mailtext%", text, -1)
	if rc == "enabled" {
		mailxml = strings.Replace(mailxml, "%rc%", TplSendRC, -1)
	} else {
		mailxml = strings.Replace(mailxml, "%rc%", "", -1)
	}
	resp, err := senditem(mailxml, "r1")
	if err != nil {
		log.Printf("Error %q\n", err)
	}
	ioutil.WriteFile("s1.log", []byte(resp), 0755)
	rx := regexp.MustCompile(`(?s)<t:ItemId\s+Id="(.*?)"\s+ChangeKey="(.*?)"`)
	tkn := rx.FindStringSubmatch(resp)
	if len(tkn) != 3 {
		log.Printf("Error tkn: %v\n", tkn)
		return "", ""
	}
	return tkn[1], tkn[2]
}

// return item-key
func SendAttachStep(msgid string, msgkey string, files []string) string {
	mailxml := TplSendAttach
	mailxml = strings.Replace(mailxml, "%mailid%", msgid, -1)
	mailxml = strings.Replace(mailxml, "%mailkey%", msgkey, -1)
	fixml := ""
	for i := 0; i < len(files); i++ {
		tmpxml := TplAttachItem
		tmpxml = strings.Replace(tmpxml, "%filename%", files[i], -1)
		tmpxml = strings.Replace(tmpxml, "%filecontent%", FileContent[files[i]], -1)
		fixml += tmpxml
	}
	mailxml = strings.Replace(mailxml, "%mailfiles%", fixml, -1)
	resp, err := senditem(mailxml, "r2")
	if err != nil {
		log.Printf("Error %q\n", err)
	}
	ioutil.WriteFile("s2.log", []byte(resp), 0755)
	rx := regexp.MustCompile(`(?s)RootItemChangeKey="(.+?)"`)
	tkn := rx.FindStringSubmatch(resp)
	if len(tkn) != 2 {
		log.Printf("Error tkn: %v\n", tkn)
		return ""
	}
	//mmid := tkn[1]
	return tkn[1]

}

func SendLetterFinal(msgid string, msgkey string) {
	mailxml := TplSendFinal
	mailxml = strings.Replace(mailxml, "%mailid%", msgid, -1)
	mailxml = strings.Replace(mailxml, "%mailkey%", msgkey, -1)
	resp, err := senditem(mailxml, "r3")
	if err != nil {
		log.Printf("Error %q\n", err)
	}
	ioutil.WriteFile("s3.log", []byte(resp), 0755)
	// s3 err handling
	s3err = !strings.Contains(resp, "NoError")
	if s3err {
		rx := regexp.MustCompile(`(?s)<m:MessageText>(.+?)</m:MessageText>`)
		ekn := rx.FindStringSubmatch(resp)
		if len(ekn) != 2 {
			lasterr = "???"
		} else {
			lasterr = ekn[1]
		}
	}
}
