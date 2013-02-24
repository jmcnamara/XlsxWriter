#!/usr/bin/perl -i

# Simple utility to modify the TeX output prior to creating a pdf file.

use strict;
use warnings;


while (<>) {

    # Convert escaped single quotes back to real single quote so that
    # the Latex upquote package has an effect.
    s/\\PYGZsq{}/'/g;


    # Modify the Pygments formatting.
    #
    # Remove italic.
    s/\\let\\PYG\@it=\\textit//g;

    # Change the comments color.
    s/0\.25,0\.50,0\.56/0.40,0.69,0.33/;


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

    # Modifiy the pre-amble. We could do this in the Sphinx conf.py
    # but ReadTheDocs doesn't support the fonts.
    if ( /^\\usepackage{sphinx}/ ) {
        print "\\usepackage{upquote}\n";
        print "\\usepackage{DejaVuSansMono}\n";
        print "\\usepackage[T1]{fontenc}\n";
        print "\\usepackage{helvet}\n";
        print "\\renewcommand{\\familydefault}{\\sfdefault}\n";
    }
}


__END__
