#/bin/bash

# xmllint.exe, docx2odt.exe, docx2txt.exe must be in PATH

if [ $# -ne 2 ]; then
    echo "usage: _docx2git.sh <infile.docx> <output folder>"
    exit 1
fi

docname="$1"
gitfolder="$2"

docname_base=${docname%%.docx}

# echo ${docname}
# echo ${docname_base}
# echo ${gitfolder}
# exit

# create repo folder
# no error if folder already exists
mkdir -p ${gitfolder}

# convert to text, fold and move to git folder
docx2txt.exe "${docname}"
dos2unix "${docname_base}".txt
echo "folding to textwidth 80 ..."
# fold -w 80 -s "${docname_base}".txt | sed '/^$/{N;/^\n$/d;}' | sed 's/  */ /g' | sed 's/ \{1,\}/ /g' | tr -s ' ' > content.txt
fold -w 80 -s "${docname_base}".txt | sed 's/\t/    /g' | cat -s | tr -s ' ' > content.txt
mv content.txt ${gitfolder}

# convert to odt, unzip it to docname_base folder
docx2odt.exe "${docname}"

# xmllint inside document_base
cd "${docname_base}"
for xml in $(ls *.xml); do
    xmllint --format ${xml} > out.xml
    mv out.xml ${xml}
done
cd ..

# copy document_base to git folder
cp -r ./"${docname_base}"/* ${gitfolder}

# cleanup
rm "${docname_base}".txt
rm -r ./"${docname_base}"

