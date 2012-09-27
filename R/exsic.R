
#mstr <- NULL
#sind <- NULL
#cntr <- NULL
#db.bn<- NULL

is.windows <- function(){
  str_detect(Sys.getenv("OS"),"Windows")
}

read.excel <-function(file, sheet, stringsAsFactors=F){
  data=NULL
  try({
  con = odbcConnectExcel(file)
  data = sqlFetch(con,sheet)
  odbcCloseAll()
  if(!stringsAsFactors){
    n = ncol(data)
    for(i in 1:n){
      if(class(data[,i])=="factor") data[,i]=as.character(data[,i])
    }
  }
  }, silent=T)
  data
}

load.data <-function(data.file){
  try({
  #mstr = get("mstr")
  #sind = get("sind")
  #cntr = get("sind")
  mstr <-read.excel(data.file,sheet="data")
  sind <-read.excel(data.file,sheet="species-order")
  cntr <-read.excel(data.file,sheet="countries")
  
  # apply filters
  if(cntr$ORIGCTY[1]!="all"){
    uctr = sort(unique(cntr$ORIGCTY))
    mstr <- mstr[mstr$ORIGCTY %in% uctr,]
  }
  mstr <- mstr[mstr$SPECIES %in% sind$SPECIES,]
  if(length(which(mstr$remove=="x"))>0) {
    mstr <- mstr[-which(mstr$remove=="x"),]
  }
  
  #Remove rows without minimal data in certain field
  if(length(which(is.na(mstr$GENUS))>0)) mstr <- mstr[-which(is.na(mstr$GENUS)),]
  if(length(which(is.na(mstr$SPECIES))>0))mstr <- mstr[-which(is.na(mstr$SPECIES)),]
  if(length(which(is.na(mstr$collector))>0))mstr <- mstr[-which(is.na(mstr$collector)),]
  if(length(which(is.na(mstr$number))>0))mstr <- mstr[-which(is.na(mstr$number)),]
  if(length(which(is.na(mstr$ORIGCTY))>0))mstr <- mstr[-which(is.na(mstr$ORIGCTY)),]
  
  cite = 1:nrow(mstr)
  spi = 1:nrow(mstr)
  for(i in 1:nrow(mstr)){
    cite[i] = get.coll.cite(mstr[i,"collector"], mstr[i,"addcoll"], mstr[i,"use.initial"])
    sp=mstr[i,"SPECIES"]
    spi[i] = sind[sind$SPECIES==sp,c("order")]
  }
  mstr <- cbind(mstr, cite, spi)
  mstr$cite <- as.character(mstr$cite)
  },silent=T)

  #assign("mstr",mstr)
  #assign("cntr",cntr)
  #assign("sind",sind)
  save(mstr,file = "mstr.Rda")
  save(cntr,file = "cntr.Rda")
  save(sind,file = "sind.Rda")
}

get.coll.cite = function(coll, addcoll, use.initial=NULL){
  coll.last = strsplit(coll,",")[[1]][1]
  res = coll.last
  addc.cnt = strsplit(addcoll,";")[[1]]
  if(length(addc.cnt)>0){
    n = length(addc.cnt)
    if(n>1){
      res=paste(res," et al.",sep="")
    }
    if(n==1){
      add.ln = strsplit(addc.cnt[1],",")[[1]][1]
      res = paste(res," & ", add.ln, sep="")
    }
  }
  if(!is.null(use.initial)){
    if("x"==use.initial){
      res=gsub(",","",coll)
    }
  }
  res = str_replace(res," & NA","")
  res
}


format.date=function(d,m=NULL,y=NULL){
  if(!is.na(d)){
  if(is.null(m) & is.null(y)){
    #then assume MCPD format definition for COLLDATE
    dmy = d
    y = str_sub(dmy, 1,4)
    m = str_sub(dmy, 5,6)
    d = str_sub(dmy, 7,8)
    if(m=="00" | m=="--") m = NA
    if(d=="00" | d=="--") d = NA
    d = as.integer(d)
    m = as.integer(m)
    y = as.integer(y)
  }
  
  mm = c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
  dd=""
  if(!is.na(d)){
    dd = paste(d,sep="")
  }
  if(!is.na(m)){
    dd = paste(dd," ",mm[m],sep="")
  } else {
    dd=""
  }
  if(!is.na(y)){
    dd = paste(dd," ",y,sep="")
  }
  } else {
    dd=""
  }
  dd
}

format.degr=function(lat,lon){
  if(is.na(lat) | is.na(lon)){
    return("")
  } else {
    latl = ""
    if(lat<0) {
      latl="S"
    } else if(lat>0){
      latl="N"
    }
    latd = abs(round(lat,2))
    latm = round(abs(floor(latd)-latd)*60,0)
    latd = floor(latd)
    lats = paste(latd,"\u00B0",latm,"'",latl,sep="")
    
    lonl = ""
    if(lon<0) {
      lonl="W"
    } else if(lon>0){
      lonl="E"
    }
    lond = abs(round(lon,2))
    lonm = round(abs(floor(lond)-lond)*60,0)
    lond = floor(lond)
    lons = paste(lond,"\u00B0",lonm,"'",lonl,sep="")
    paste(lats,", ",lons,sep="")
  }
}

format.alt=function(alt){
  res=""
  if(!is.na(alt)) {
    res=paste(alt," m",sep="")
  }
  res
}


format.coll=function(coll, num){
  res=paste(coll," ",num,sep="")
  res
}

format.dups=function(dups){
  res=""
  if(!is.na(dups)){
    #sort dups!
    dps=paste(" (",dups,")",sep="")
    if(dps==" ()") dps=""
    res = paste(res,dps, sep="")
  }
  res
}


get.record=function(i, data1, ui=NULL){
  if(is.null(ui)){
    r = data1[data1$id==i,
              c("collector","addcoll","number","COLLDATE",
                "ELEVATION","COLLSITE","DECLATITUDE","DECLONGITUDE","dups")]
    r = r[1,]
    cl = get.coll.cite(r[["collector"]], r[["addcoll"]])
  } else {
    r = data1[data1$id==i,
              c("use.initial","collector","addcoll","number","COLLDATE",
                "ELEVATION","COLLSITE","DECLATITUDE","DECLONGITUDE","dups")]
    cl = get.coll.cite(r[["collector"]], r[["addcoll"]],r[["use.initial"]])
  }
  dt = format.date(r[["COLLDATE"]])
  dg = format.degr(r[["DECLATITUDE"]],r[["DECLONGITUDE"]])
  da = format.alt(r[["ELEVATION"]])
  
  
  dc = format.coll(cl,r[["number"]])
  dp = format.dups(r[["dups"]])
  
  rec1=paste(r["COLLSITE"],dg,da,dt,sep=", ")
  rec1=gsub(" ,","",rec1)
  
  rec1=paste(rec1,", *",dc,sep="")
  rec1=paste(rec1,"* ",str_trim(dp),"",sep="")
  rec1
}


md2html <-function(file){
  file = str_replace(file,".Rmd",".md")
  txt = readLines(file)
  txt = str_replace(txt,'`','')
  writeLines(txt,file)
  fn = str_split(file,"\\.")[[1]][1]
  fh = paste(fn,".html",sep="")
  markdownToHTML(file,fh)
  txt = readLines(fh)
  txt = str_replace(txt,'utf-8','utf-64')
  writeLines(txt,fh)
  
}

index.countries <-function # Creates a simple index of countries.
### Creates a simple index of countries as text string.
(){
##note<< Should only be used within the template file.
  #cntr = get("cntr")
  #mstr = get("mstr")
  cntr = NULL
  mstr = NULL
  load("cntr.Rda")
  load("mstr.Rda")
  if(cntr$ORIGCTY[1]=="all"){
    cn = sort(unique(mstr$ORIGCTY))
  } else {
    cn = sort(unique(cntr$ORIGCTY))  
  }
  
  n = length(cn)
  res=paste(1:n,". ",cn,sep="")
  paste(res,collapse="\n")
### A string containing the species formatted 
### for the template file in markdown markup.
}

generic.list <-function
(cl="\n") {
  #sind = get("sind")
  sind = NULL
  load("sind.Rda")
  sp = paste(". *",substr(sind$GENUS,1,1),". ",sind$SPECIES,"*",sep="")
  n = 1:length(sp)
  x=paste(n,sp, sep="")
  paste(x,collapse=cl)
}


index.species <-function # Creates a simple index of species.
### Creates an index of species as text string.
  ##note<< Should only be used within the template file.
  () {
  generic.list()
  ### A string containing the species formatted 
  ### for the template file in markdown markup.
}

index.collections <-function # Creates the index of collectors and their specimes.
### Creates the index of collectors and their specimes.
  ##note<< Should only be used within the template file.
(){
  res=""
  #mstr = get("mstr")
  mstr = NULL
  load("mstr.Rda")

  if(nrow(mstr)!=0){
  col.ord = sort(unique(mstr$cite))
  #srt = c("cite","number")
  #dbs = mstr[do.call(order,srt),]
  dbs = arrange(mstr,mstr$cite, mstr$number)
  
  for(i in 1:length(col.ord)){
    res = paste(res,col.ord[i]," ",sep="")
    dbt = dbs[dbs$cite==col.ord[i],]
    
    for(k in 1:nrow(dbt)){
      res = paste(res,dbt$number[k]," (",dbt$spi[k],")",sep="")
      if (k<nrow(dbt)){
        res = paste(res,", ",sep="")
      } else {
        res = paste(res,".",sep="")
      }
    }
    res=paste(res,'\n')
  }
  } else {
    res = "**No records found for this combination of countries and species.**\n"
  }
  res
  ### A string containing the collectors formatted 
  ### for the template file in markdown markup.
  
}

index.species.short<-function # A condensed list of species (listed within a line).
### A condensed list of species (listed within a line).
  ##note<< Should only be used within the template file.
  (){
  paste("[",generic.list(", "),"]",sep="")
  ### A string containing the species formatted 
  ### for the template file in markdown markup.
}

index.specimens <-function # The list of specimens 
 ### List of specimens by species and country detailing collector,
 ### collector number, date and location.
 ##note<< Should only be used within the template file.
  (){
  #sind = get("sind")
  #mstr = get("mstr")
  sind = NULL 
  load("sind.Rda")
  mstr = NULL
  load("mstr.Rda")
  if(nrow(mstr)==0){
    res = "**No records found for this combination of countries and species.**\n"
  } else {
  res=""  
  for(i in 1:nrow(sind)){
    
    sp = sind[i,]$SPECIES
    db = mstr[mstr$SPECIES==sp,]
    res=paste(res,"### ",i,". ",sind$GENUS[i]," ",sp,"\n",sep="")
    if(nrow(db)==0){
      res = paste(res,
                  "**No records for this species found in this database.**\n"
                  ,sep="")
    } else {
    #sort by collector last name & number
    srt = c("ORIGCTY","majorarea","collector","number")
    dbs = db[do.call(order,db[srt]),]
    
      
    #filter by countries
    fcntry = sort(unique(dbs$ORIGCTY))
    for(j in 1:length(fcntry)){
      if(nrow(dbs)==0){
        res = paste(res,"**No records for this species in ", fcntry[1]
                    ," in this database.**\n",sep="")
      } else {
        
      db1 = dbs[dbs$ORIGCTY==fcntry[j],]
      res=paste(res,"**",fcntry[j],".** ",sep="")
      
      #filter by admin1
      fadm1 = sort(unique(db1$majorarea))
      for(k in 1:length(fadm1)){
        res = paste(res, toupper(fadm1[k]),": ",sep="")
        db2 = db1[db1$majorarea==fadm1[k],]
        db2 = db2[!is.na(db2$id),]
        for(l in 1:nrow(db2)){
          res=paste(res,get.record(db2[l,"id"],db2),sep="")
          if(l<nrow(db2)){
            final = "; "
          } else {
            final = "."
          }
          res=paste(res,final,sep="")
        }
        if(k<length(fadm1)){
          res=paste(res," - ",sep="")
        }
      
      }
      res=paste(res,"\n\n",sep="")
      }
      }
    
      
    }
    
  }
  }
  res
  ### A string containing the specimen details formatted 
  ### for the template file in markdown markup.
}

exsic.database <- function # Get hyperlink to database of passport data
### Creates a hyperlink in markdown format.
(){
  #db.bn = get("db.bn")
  db.bn = NULL
  load("db.bn.Rda")
  paste("[Database file](file:///",file.path(getwd(),db.bn),")",sep="")
  ### A string containing the hyperlink to the file in markdown format.
}

update.pb <- function(pb, i, txt){
  info <- paste(sprintf("%d%% done", round(i))," (",txt,")",sep="")
  setWinProgressBar(pb, i, "Progess in percent:", info)
}


exsic <-structure(function # Create specimen indices.
### Creates four indices based on passport data of a database 
### of biological specimens.
(
  data.file=NULL, ##<< Path to the Excel file containing the database
  template.file=system.file("templates/exsic-simple.Rmd",package='mcpd') ##<< Path to the template file
  ){
  if(is.null(data.file)) {
    data.file = file.choose()
  }
  pb=NULL
  if(interactive() & is.windows()) {
	pb <- winProgressBar("Creating botanical indices", "Progress in %",
                       0, 100, 0)
  }

  if(interactive() & is.windows()) {update.pb(pb,10, "Analyzing database.")}
  fr = template.file
  bn = basename(fr)
  tf.to = file.path(getwd(),bn)
  if(!file.exists(tf.to)) file.copy(fr, tf.to, overwrite=T)
  
  fr = data.file
  db.bn <- basename(fr)

  save(db.bn, file = "db.bn.Rda")
  db.to = file.path(getwd(),db.bn)
  if(!file.exists(db.to)) file.copy(fr, db.to, overwrite=T)
  
  ##note<< A better check routine for compliance with MCPD data fields is pending.
  # TODO check file format & content
  # TODO check.data.file(data.file)
  kn.to = file.path(getwd(),db.bn)
  kn.to = str_replace(kn.to,".xls",".md")
  hm.to = str_replace(kn.to,".md",".html")
  
  unlink(hm.to)
  if(interactive() & is.windows()) {update.pb(pb,20, "Loading data.")}
  load.data(db.to)  
  mstr=NULL
  load("mstr.Rda")
  if(!is.null(mstr)){
    if(interactive() & is.windows()) {update.pb(pb,40, "Transforming data.")}
    knit(tf.to, kn.to )
    if(interactive() & is.windows()) {update.pb(pb,80, "Formatting web page.")}
    md2html(kn.to)
    if(interactive() & is.windows()) {update.pb(pb,100, "Finalizing.")}
  } else {
    unlink(hm.to)
    aline = paste("<html><body><p><b>Database file ",db.bn," could not be read.
                  </b></p></body></html>")
    write(aline, hm.to)
  }
  unlink("mstr.Rda")
  unlink("sind.Rda")
  unlink("cntr.Rda")
  unlink("db.bn.Rda")
  if(interactive() & is.windows()) {close(pb)}
  if(file.exists(hm.to) & interactive() & is.windows()) {shell.exec(hm.to)}
  else {
    print(paste("Results are in file:",hm.to))
  }
},ex=function(){
  exsic(system.file("example/specimen-database.xls",package="mcpd"))
})





