#!/usr/bin/perl -i

# Simple utility to modify the TeX output prior to creating a pdf file.

use strict;
use warnings;


while (<>) {

    # Change scale of images and center them.
    if ( s/^\\includegraphics/\\includegraphics[scale=0.75]/ ) {
        print "\\begin{center}\n";
        print;
        print "\\end{center}\n";

        next;
    }

    # Wrap Verbatim sections in "quote" to indent.
    if ( /^\\begin{Verbatim}/ ) {
        print "\\begin{quote}\n";
    }

    print;

    if ( /^\\end{Verbatim}/ ) {
        print "\\end{quote}\n";
    }
}


__END__
