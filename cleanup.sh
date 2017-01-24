#!/bin/bash

script=`readlink -e $0`
scriptname=`basename $script`
ooxxsl=`dirname $script`

# command-line options
debug=1
help=1

usage () {
    echo "$scriptname [-h|--help] [-d|--debug] docxfile"
    return 0
}

if [ $# -lt 1 -o $# -gt 2 ] ; then
    help=0
else
    while (( $# )) ; do
        case "$1" in
            -h | --help)
                help=0
                ;;
            -d | --debug)
                debug=0
                ;;
            -*)
                help=0
                ;;
            *)
                infile=$1
                ;;
        esac
        shift
    done
fi

if [ "$help" == "0" ] ; then
    usage
    exit 1
fi

if [ ! \( -f $infile -a -r $infile \) ] ; then
    echo $scriptname": input file "$infile" not a readable file" 2>&1
    exit 1
fi

project=`basename $infile '.docx'`
pid=$$
infile=`readlink -e $infile`

clean_content () {
    source=$1
    target=$2
    stage=$3
    cp -r $source $target
    for ooxfile in word/document.xml word/endnotes.xml word/footnotes.xml ; do
        if [ -f $source/$ooxfile ] ; then
            rm -f $target/$ooxfile
            xsltproc $ooxxsl/$stage"_clean.xsl" $source/$ooxfile > $target/$ooxfile
            xsltproc=$?
            if [ "$xsltproc" != "0" ] ; then
                echo $scriptname": xsltproc failed at stage"$stage 2>&1
                return $xsltproc
            fi
        fi
    done
    return 0
}

zip_content() {
    target=$1
    pushd $target > /dev/null
    zip -rD ../$target".docx" \[Content_Types\].xml _rels word customXml docProps > /dev/null
    zip=$?
    popd > /dev/null
    return $zip
}

# unpack the original
source=$project"_"$pid"_orig"
mkdir $source
pushd $source > /dev/null
unzip $infile > /dev/null
unzip=$?
if [ "$unzip" != "0" ] ; then
    echo $scriptname": can’t unzip input file "$infile 2>&1
    popd > /dev/null
    exit $unzip
fi
popd > /dev/null

previous=$source
#for stage in fonts normalize punctuation ; do
for stage in fonts strip normalize ; do
    target=$project"_"$pid"_"$stage
    clean_content $previous $target $stage
    result=$?
    if [ "$result" != "0" ] ; then
        exit $result
    fi
    if [ "$debug" == "0" ] ; then
        zip_content $target
        zip=$?
        if [ "$zip" != "0" ] ; then
            echo $scriptname": zipping intermediate directory "$target" failed" 2>&1
            exit $zip
        fi
    else
        rm -rf $previous
    fi
    previous=$target
done

# clean up style usage in the document
style=$project"_"$pid"_style"
cp -r $previous $style
for ooxfile in word/document.xml word/endnotes.xml word/footnotes.xml ; do
    if [ ! -f $previous/$ooxfile ] ; then
        echo "<?xml version=\"1.0\"?><a/>" > $previous/$ooxfile
    fi
done
ooxfile=word/styles.xml
if [ -f $previous/$ooxfile ] ; then
    rm -f $style/$ooxfile
    xsltproc $ooxxsl/style_clean.xsl $previous/$ooxfile > $style/$ooxfile
    xsltproc=$?
    if [ "$xsltproc" != "0" ] ; then
        echo $scriptname": xsltproc failed" 2>&1
        exit $xsltproc
    fi
fi
zip_content $style
zip=$?
if [ "$zip" != "0" ] ; then
    echo $scriptname": zipping intermediate directory "$style" failed" 2>&1
    exit $zip
fi
if [ "$debug" != "0" ] ; then
    rm -rf $previous $style
    mv -i $style".docx" $project"_out.docx"
fi

exit 0
