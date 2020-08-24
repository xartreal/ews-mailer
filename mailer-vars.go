// mailer-vars
package main

var (
	from     string
	endpoint string
	username string
	userpass string
)

//attach list
var (
	filenames []string
	filecount int
)

// attach content
var FileContent = map[string]string{}

var (
	s3err   bool
	lasterr string
)

