\name{mcpd-package}
\alias{mcpd-package}
\docType{package}
\title{Tools to analyse and use passport data for biological collections.}
\description{The package provides some tools for botanists, taxonomists, curators of plant genebanks and perhaps other biological collections that make use of passport data for plant specimens in a standard format (MCPD). Currently, the tool allows to create simple specimen indices as found in taxonomic treatments. An example Excel file is provided. Each specimen record must at least have content in the following fields: GENUS, SPECIES, collector, number, ORIGCTY (names in upper case correspond to fields defined in the MCPD standard.). In addition, records can be removed explicitly from the generated index putting an 'x' in the 'remove' field. Further filtering can be done by specifying countries in the 'countries'-sheet or allowing 'all'. Countries will be ordered alphabetically. Species can be ordered and filtered explicitly in the sheet 'species-order'. The geographic coordinate fields 'DECLATITUDE' and 'DECLONGITUDE' are assumed to be decimal and in WGS84 projection. As an output, a web page in standard html format is created based on a template in markdown format. The template may be changed and specified. The generated web-page can be edited in most word processing software.}
\details{
\tabular{ll}{Package: \tab mcpd\cr
Type: \tab Package\cr
Title: \tab Tools to analyse and use passport data for biological collections.\cr
Version: \tab 1.0\cr
Date: \tab 2012-09-27\cr
Author: \tab Reinhard Simon\cr
Maintainer: \tab Reinhard Simon <r.simon@cgiar.org>\cr
License: \tab GPL 2.0 or higher\cr
Depends: \tab 
RODBC,
knitr,
markdown,
stringr,
plyr\cr}
}
\author{Reinhard Simon}

\keyword{ package }


