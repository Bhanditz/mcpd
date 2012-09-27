\name{exsic}
\alias{exsic}
\title{Create specimen indices.}
\description{Creates four indices based on passport data of a database 
of biological specimens.}
\usage{exsic(data.file = NULL, template.file = system.file("templates/exsic-simple.Rmd", 
    package = "mcpd"))}
\arguments{
  \item{data.file}{Path to the Excel file containing the database}
  \item{template.file}{Path to the template file}
}



\author{Reinhard Simon}
\note{A better check routine for compliance with MCPD data fields is pending.}



\examples{
exsic(system.file("example/specimen-database.xls",package="mcpd"))
}
