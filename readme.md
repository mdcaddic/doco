This doco repository:

1. enables testing using markdown documents as sections of a greater document and then converting that to docx format using pandoc
2. enables testing the automatic generation of ABAC (As Built As Configured) documents and the conversion from markdown to word using pandoc


echo "# freqtrade" >> README.md

git init

git add README.md

git commit -m "first commit"

git branch -M main

git remote add origin https://github.com/mdcaddic/freqtrade.git

git push -u origin main
