#!/bin/bash

clear
echo "|"
echo "| Pre-release checks."
echo "|"
echo


#############################################################
#
# Run tests.
#
function check_test_status {

    echo
    echo -n "Are all tests passing for all Pythons? [y/N]: "
    read RESPONSE

    if [ "$RESPONSE" != "y" ]; then

        echo -n "    Run all tests now?                 [y/N]: "
        read RESPONSE

        if [ "$RESPONSE" != "y" ]; then
            echo
            echo -e "Please run: make testpythonsall\n";
            exit 1
        else
            echo "    Running tests...";
            make testpythonsall
            check_test_status
         fi
    fi
    clear
}


#############################################################
#
# Run spellcheck.
#
function check_spellcheck {

    echo
    echo -n "Is the spellcheck ok?                  [y/N]: "
    read RESPONSE

    if [ "$RESPONSE" != "y" ]; then

        echo -n "    Run spellcheck now?                [y/N]: "
        read RESPONSE

        if [ "$RESPONSE" != "y" ]; then
            echo
            echo -e "Please run: make spellcheck\n";
            exit 1
        else
            echo "    Running spellcheck...";
            make spellcheck
            check_spellcheck
         fi
    fi
    clear
}


#############################################################
#
# Run testpep8.
#
function check_testpep8 {

    echo
    echo -n "Is the testpep8 ok?                  [y/N]: "
    read RESPONSE

    if [ "$RESPONSE" != "y" ]; then

        echo -n "    Run testpep8 now?                [y/N]: "
        read RESPONSE

        if [ "$RESPONSE" != "y" ]; then
            echo
            echo -e "Please run: make testpep8\n";
            exit 1
        else
            echo "    Running testpep8...";
            make testpep8
            check_testpep8
         fi
    fi
}


#############################################################
#
# Check Changes file is up to date.
#
function check_changefile {
    clear

    echo "Latest change in Changes file: "
    perl -ne '$rev++ if /^Release/; exit if $rev > 1; print "    | $_"' Changes

    echo
    echo -n "Is the Changes file updated? [y/N]: "
    read RESPONSE

    if [ "$RESPONSE" != "y" ]; then
        echo
        echo -e "Please update the Change file to proceed.\n";
        exit 1
    fi
}


#############################################################
#
# Check the versions are up to date.
#
function check_versions {

    clear
    echo
    echo "Latest file versions: "

    grep -He "[0-9]\.[0-9]\.[0-9]" setup.py dev/docs/source/conf.py xlsxwriter/__init__.py | sed 's/:/ : /g' | sed 's/=/ = /' | awk '{printf "    | %-24s %s\n", $1, $5}'

    echo
    echo -n "Are the versions up to date?   [y/N]: "
    read RESPONSE

    if [ "$RESPONSE" != "y" ]; then
        echo -n "    Update versions?           [y/N]: "
        read RESPONSE

        if [ "$RESPONSE" != "y" ]; then
            echo
            echo -e "Please update the versions to proceed.\n";
            exit 1
        else
            echo "    Updating versions...";
            perl -i dev/release/update_revison.pl setup.py dev/docs/source/conf.py xlsxwriter/__init__.py
            check_versions
         fi
    fi
}

#############################################################
#
# Check that the PDF doc is up to date.
#
function check_pdf_doc {

    clear
    echo
    echo -n     "Is the PDF doc up to date?   [y/N]: "
    read RESPONSE

    if [ "$RESPONSE" != "y" ]; then
        echo -n "    Rebuild pdf?             [y/N]: "
        read RESPONSE

        if [ "$RESPONSE" == "y" ]; then
            make -C dev/docs latexpdf
            cp -r dev/docs/build/latex/XlsxWriter.pdf docs
        fi
    fi
}

#############################################################
#
# Run release checks.
#
function check_git_status {
    clear

    echo "Git status: "
    git status | awk '{print "    | ", $0}'

    echo "Git log: "
    git log -1 | awk '{print "    | ", $0}'

    echo "Git latest tag: "
    git tag -l -n1 | tail -1 | awk '{print "    | ", $0}'

    echo
    echo -n "Is the git status okay? [y/N]: "
    read RESPONSE

    if [ "$RESPONSE" != "y" ]; then
        echo
        echo -e "Please fix git status.\n";

        git tag -l -n1 | tail -1 | perl -lane 'printf "git commit -m \"Prep for release %s\"\ngit tag \"%s\"\n\n", $F[4], $F[0]' | perl dev/release/update_revison.pl
        exit 1
    fi
}

check_test_status
check_spellcheck
check_testpep8
check_changefile
check_versions
check_pdf_doc
check_git_status


#############################################################
#
# All checks complete.
#
clear
echo
echo "Interface configured [OK]"
echo "Versions updated     [OK]"
echo "Git status           [OK]"
echo
echo "Everything is configured.";
echo

echo -n "Confirm release: [y/N]: ";
read RESPONSE

if [ "$RESPONSE" == "y" ]; then
    exit 0
else
    exit 1
fi
