# outlook-addin-sample-boilerplate

## Docs and Links
### Manifest Official Docs
https://docs.microsoft.com/en-us/outlook/add-ins/manifests

### Sample Manifest
https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml

## Notes
Url used: https://localhost/dist/page-main.html


## Generate Self-Signed Certs
https://github.com/synle/node-proxy-example
https://unix.stackexchange.com/questions/90450/adding-a-self-signed-certificate-to-the-trusted-list
https://askubuntu.com/questions/73287/how-do-i-install-a-root-certificate


```
openssl genrsa -des3 -passout pass:x -out server.pass.key 2048
openssl rsa -passin pass:x -in server.pass.key -out server.key
rm server.pass.key
openssl req -new -key server.key -out server.csr
openssl x509 -req -sha256 -days 365 -in server.csr -signkey server.key -out server.crt

sudo dpkg-reconfigure ca-certificates
```


