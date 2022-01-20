#/bin/bash

if [ $# -ne 2 ]; then
    echo "usage: _docx_repo_update.sh <infile.docx> <output folder>"
    exit 1
fi

docname="$1"
gitfolder="$2"

zipname=${docname%%.docx}

# echo ${docname}
# echo ${zipname}
# echo ${gitfolder}
# exit

# create repo folder
# no error if folder already exists
mkdir -p ${gitfolder}

# must be in PATH
docx2odt.exe "${docname}"

cd "${zipname}"
for xml in $(ls *.xml); do
    xmllint --format ${xml} > out.xml
    mv out.xml ${xml}
done
cd ..

cp -r ./"${zipname}"/* ${gitfolder}

