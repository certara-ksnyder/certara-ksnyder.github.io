# app.R  — Shinylive-safe Excel export (browser only, no downloadHandler)
library(shiny)
library(jsonlite)
library(bslib)

options(shiny.maxRequestSize = 200 * 1024^2)

`%||%` <- function(a, b) if (is.null(a)) b else a
bind_rows_base <- function(a, b) {
  if (is.null(a)) return(b)
  cols <- union(names(a), names(b))
  for (nm in setdiff(cols, names(a))) a[[nm]] <- NA
  for (nm in setdiff(cols, names(b))) b[[nm]] <- NA
  a <- a[, cols, drop = FALSE]; b <- b[, cols, drop = FALSE]
  rbind(a, b)
}
from_excel_date <- function(x) {
  suppressWarnings({
    xnum <- suppressWarnings(as.numeric(x))
    d <- as.Date(xnum, origin = "1899-12-30")
    d[is.na(d)] <- as.Date(x[is.na(d)])
    d
  })
}

ui <- page_fillable(
  theme = bs_theme(bootswatch = "flatly"),
  title = "SEND Mapper (Shinylive)",
  tags$head(
    # Load SheetJS from local /www so it works offline and on Pages
    tags$script(src = "xlsx.full.min.js"),
    # JS bridge to parse & export
    tags$script(src = "custom.js"),
    tags$style(HTML(".req{color:#a94442;font-weight:600}.ok{color:#2e7d32} code,pre{font-size:.9rem}"))
  ),
  layout_sidebar(
    sidebar = sidebar(
      h4("1) Upload legacy file(s)"),
      fileInput("legacy_files", "(.xlsx / .xls / .csv)", accept = c(".xlsx",".xls",".csv"), multiple = TRUE),
      helpText("Shinylive parses in the browser, classic Shiny can read on the server."),
      hr(),
      h4("2) Study / Template"),
      textInput("studyid", "STUDYID", placeholder = "e.g., 405775-20250206-CPK"),
      textInput("usubid_col_prefix", "USUBJID prefix (default = STUDYID)"),
      dateInput("rfstdtc", "DM.RFSTDTC (Start)"),
      dateInput("rfendtc", "DM.RFENDTC (End)"),
      hr(),
      actionButton("reset_all", "Reset", class = "btn btn-outline-danger")
    ),
    card(
      card_header("Workflow"),
      tabsetPanel(
        id = "nav",
        tabPanel("A. Select Sheets & Preview",
                 fluidRow(
                   column(4,
                          uiOutput("sheet_picker"),
                          checkboxInput("first_row_header","First row is header", TRUE),
                          checkboxInput("has_units_row","Second row has UNITS (CRO-1)", FALSE),
                          checkboxInput("wide_tests","Tests are columns (wide)", TRUE),
                          actionButton("parse_now","Parse / Refresh Preview", class="btn btn-primary")
                   ),
                   column(8, h5("Preview (first 25 rows)"), tableOutput("preview"))
                 )
        ),
        tabPanel("B. Map Columns",
                 fluidRow(
                   column(4,
                          uiOutput("id_cols_ui"), hr(),
                          selectInput("specimen","Default LBSPEC", c("SERUM","PLASMA","WHOLE BLOOD","URINE","OTHER"), "SERUM"),
                          selectInput("category","Default LBCAT", c("CLINICAL CHEMISTRY","HEMATOLOGY","COAGULATION","URINALYSIS"), "CLINICAL CHEMISTRY"),
                          checkboxInput("map_urine_semiquant","Map UA +/- to STRESN (0..4)", TRUE),
                          actionButton("auto_map","Auto-map common analytes", class="btn btn-secondary")
                   ),
                   column(8, h5("Analyte Mapping (LBTESTCD → source column)"), uiOutput("analyte_map_ui"))
                 )
        ),
        tabPanel("C. Metadata (DM / TS / TA / TX)",
                 fluidRow(
                   column(6, h5("DM"),
                          textInput("dm_sex_default","Default DM.SEX"),
                          textInput("dm_armcd_default","Default DM.ARMCD"),
                          textInput("dm_setcd_default","Default DM.SETCD")
                   ),
                   column(6, h5("TS (key)"),
                          textInput("ts_species","SPECIES"),
                          textInput("ts_strain","STRAIN"),
                          textInput("ts_route","ROUTE"),
                          textInput("ts_sndigver","SNDIGVER"),
                          textInput("ts_sndctver","SNDCTVER"),
                          textInput("ts_title","STITLE")
                   )
                 ),
                 hr(),
                 fluidRow(
                   column(6, h5("TA"), textInput("ta_armcd","ARMCD"), textInput("ta_arm","ARM (label)")),
                   column(6, h5("TX"),
                          textInput("tx_setcd","SETCD"), textInput("tx_set","SET (label)"),
                          textInput("tx_trtdos","TRTDOS"), textInput("tx_trtdosu","TRTDOSU")
                   )
                 )
        ),
        tabPanel("D. Validate & Export",
                 h5("Checks"), verbatimTextOutput("checks"), hr(),
                 h5("Export"),
                 p("Exports Excel with LB, DM, TS, TA, TX (abbreviated, for analysis/visualization)."),
                 # Always show the browser export button in Shinylive & classic Shiny
                 actionButton("export_browser", "Export Excel (.xlsx)", class = "btn btn-success"),
                 br(), br(),
                 strong("Note:"), p("Not for regulatory submissions.")
        )
      )
    )
  )
)

server <- function(input, output, session) {
  
  rv <- reactiveValues(parsed=NULL, data=NULL, cols=character(),
                       analyte_map=list(), animal_col=NULL, date_col=NULL,
                       timepoint_col=NULL, sex_col=NULL, armcd_col=NULL, setcd_col=NULL)
  
  observeEvent(input$reset_all, {
    rv$parsed <- rv$data <- NULL; rv$cols <- character(); rv$analyte_map <- list()
    rv$animal_col <- rv$date_col <- rv$timepoint_col <- rv$sex_col <- rv$armcd_col <- rv$setcd_col <- NULL
    updateTabsetPanel(session, "nav", "A. Select Sheets & Preview")
    showNotification("Reset complete", type="message")
  })
  
  # Browser-parsed files from custom.js (Shinylive path)
  observeEvent(input$excel_parsed, {
    rv$parsed <- input$excel_parsed
    showNotification(sprintf("Parsed %d file(s) in browser", length(rv$parsed$files %||% list())), type="message")
  }, ignoreInit = TRUE)
  
  # Server parsing fallback (classic Shiny run)
  observeEvent(input$legacy_files, {
    req(input$legacy_files)
    files_df <- input$legacy_files
    files <- vector("list", nrow(files_df))
    for (i in seq_len(nrow(files_df))) {
      name <- files_df$name[i]; path <- files_df$datapath[i]; ext <- tolower(tools::file_ext(name))
      sheets_list <- list()
      if (ext %in% c("xlsx","xls")) {
        if (!requireNamespace("readxl", quietly = TRUE)) {
          showNotification("Install 'readxl' for server Excel parsing: install.packages('readxl')", type="error", duration=7)
          next
        }
        sh_names <- readxl::excel_sheets(path)
        for (sn in sh_names) {
          df <- suppressWarnings(readxl::read_excel(path, sheet = sn, col_names = FALSE))
          if (!is.data.frame(df)) next
          m <- as.matrix(df); m[is.na(m)] <- ""
          rows <- lapply(seq_len(nrow(m)), function(r) as.character(m[r, ]))
          sheets_list[[length(sheets_list)+1]] <- list(name = sn, data = rows)
        }
      } else if (ext == "csv") {
        df <- try(utils::read.csv(path, header = FALSE, stringsAsFactors = FALSE, check.names = FALSE), silent=TRUE)
        if (!inherits(df, "try-error") && nrow(df)) {
          m <- as.matrix(df); m[is.na(m)] <- ""
          rows <- lapply(seq_len(nrow(m)), function(r) as.character(m[r, ]))
          sheets_list[[1]] <- list(name = tools::file_path_sans_ext(name), data = rows)
        }
      } else {
        showNotification(sprintf("Unsupported file type: %s", name), type="warning")
        next
      }
      files[[i]] <- list(name = name, sheets = sheets_list)
    }
    files <- Filter(function(x) !is.null(x) && length(x$sheets), files)
    if (length(files)) {
      rv$parsed <- list(files = files)
      showNotification(sprintf("Parsed %d file(s) on server", length(files)), type="message")
    }
  }, ignoreInit = TRUE)
  
  output$sheet_picker <- renderUI({
    req(rv$parsed)
    fl <- rv$parsed$files %||% list()
    if (!length(fl)) return(helpText("Upload files to begin."))
    ui <- tagList()
    for (i in seq_along(fl)) {
      ui <- tagAppendChildren(ui,
                              h6(sprintf("File: %s", fl[[i]]$name)),
                              selectInput(paste0("sheet_", i), NULL, choices = vapply(fl[[i]]$sheets, `[[`, "", "name"))
      )
    }
    ui
  })
  
  observeEvent(input$parse_now, {
    req(rv$parsed)
    fl <- rv$parsed$files %||% list()
    out <- NULL
    for (i in seq_along(fl)) {
      sh_name <- input[[paste0("sheet_", i)]]
      if (is.null(sh_name)) next
      sh_idx <- which(vapply(fl[[i]]$sheets, function(s) identical(s$name, sh_name), logical(1)))
      if (!length(sh_idx)) next
      rows <- fl[[i]]$sheets[[sh_idx]]$data
      if (is.null(rows) || !length(rows)) next
      
      df <- try({
        maxlen <- max(vapply(rows, length, integer(1)))
        mat <- do.call(rbind, lapply(rows, function(r) { length(r) <- maxlen; unlist(r) }))
        m <- as.data.frame(mat, stringsAsFactors = FALSE)
        
        names(m) <- if (isTRUE(input$first_row_header)) {
          hdr <- as.character(unlist(m[1,], use.names = FALSE))
          hdr[is.na(hdr) | hdr==""] <- paste0("V", which(is.na(hdr) | hdr==""))
          make.names(hdr, unique = TRUE)
        } else paste0("V", seq_len(ncol(m)))
        
        if (isTRUE(input$first_row_header)) m <- m[-1,,drop=FALSE]
        if (isTRUE(input$has_units_row)) { attr(m, "units_row") <- as.list(m[1,,drop=TRUE]); m <- m[-1,,drop=FALSE] }
        head(m, 25)
      }, silent=TRUE)
      
      if (!inherits(df, "try-error")) {
        df$..file.. <- fl[[i]]$name
        df$..sheet.. <- sh_name
        out <- bind_rows_base(out, df)
      }
    }
    rv$data <- out
    rv$cols <- if (!is.null(out)) setdiff(names(out), c("..file..","..sheet..")) else character()
    updateTabsetPanel(session, "nav", "A. Select Sheets & Preview")
  })
  
  output$preview <- renderTable(rv$data)
  
  output$id_cols_ui <- renderUI({
    req(rv$cols)
    tagList(
      selectInput("animal_col","Animal ID column", choices=c("", rv$cols)),
      selectInput("date_col","Collection date column (if present)", choices=c("", rv$cols)),
      selectInput("timepoint_col","Timepoint/visit column (optional)", choices=c("", rv$cols)),
      selectInput("sex_col","Sex column (optional)", choices=c("", rv$cols)),
      selectInput("armcd_col","Group/ARMCD column (optional)", choices=c("", rv$cols)),
      selectInput("setcd_col","SETCD column (optional)", choices=c("", rv$cols))
    )
  })
  observe({ if (isTruthy(input$animal_col))    rv$animal_col    <- input$animal_col })
  observe({ if (isTruthy(input$date_col))      rv$date_col      <- input$date_col })
  observe({ if (isTruthy(input$timepoint_col)) rv$timepoint_col <- input$timepoint_col })
  observe({ if (isTruthy(input$sex_col))       rv$sex_col       <- input$sex_col })
  observe({ if (isTruthy(input$armcd_col))     rv$armcd_col     <- input$armcd_col })
  observe({ if (isTruthy(input$setcd_col))     rv$setcd_col     <- input$setcd_col })
  
  analyte_syn <- list(
    ALB=c("ALB","Albumin"), TP=c("TP","Total Protein","T.P"), ALP=c("ALP","ALKP","Alkaline Phosphatase"),
    ALT=c("ALT","SGPT"), AST=c("AST","SGOT"), CK=c("CK","CPK"), TBIL=c("TBIL","Total Bilirubin"),
    DBIL=c("DBIL","Direct Bilirubin"), IBIL=c("IBIL","Indirect Bilirubin"), BUN=c("BUN","UREA","Urea"),
    CREAT=c("CREAT","CRE","CR"), CHOL=c("CHOL","TCHO"), GLU=c("GLU","Glucose"), PHOS=c("PHOS","P"),
    'NA'=c("NA","Na"), K=c("K","K+"), CL=c("CL"), CA=c("CA"), TG=c("TG"), GGT=c("GGT"),
    WBC=c("WBC"), NEUT=c("NEUT","NE%"), `#NEUT`=c("#NEUT","NE#"),
    LYMP=c("LYMP","LY%"), `#LYMP`=c("#LYMP","LY#"),
    MONO=c("MONO","MO%"), `#MONO`=c("#MONO","MO#"),
    EOS=c("EOS","EO%"),  `#EOS`=c("#EOS","EO#"),
    BASO=c("BASO","BA%"), `#BASO`=c("#BASO","BA#"),
    RBC=c("RBC"), HGB=c("HGB"), HCT=c("HCT"), MCV=c("MCV"), MCH=c("MCH"), MCHC=c("MCHC"),
    RDW=c("RDW"), PLT=c("PLT"), MPV=c("MPV"), `RET%`=c("%RET","RET%"), `#RET`=c("#RET","RET#"),
    PT=c("PT"), APTT=c("APTT"), FIB=c("FIB"),
    U_PH=c("PH"), U_PRO=c("PRO"), U_SG=c("S.G","SG","Specific Gravity"), U_GLU=c("GLU"), U_TURB=c("TURB")
  )
  output$analyte_map_ui <- renderUI({
    req(rv$cols)
    tgt <- sort(unique(c(
      "ALB","TP","ALP","ALT","AST","CK","TBIL","DBIL","IBIL","BUN","CREAT","CHOL","GLU",
      "PHOS","NA","K","CL","CA","TG","GGT","WBC","NEUT","#NEUT","LYMP","#LYMP","MONO","#MONO",
      "EOS","#EOS","BASO","#BASO","RBC","HGB","HCT","MCV","MCH","MCHC","RDW","PLT","MPV",
      "RET%","#RET","PT","APTT","FIB","U_PH","U_PRO","U_SG","U_GLU","U_TURB"
    )))
    do.call(tagList, lapply(tgt, function(a) {
      selectInput(paste0("map_", a), a, choices = c("", rv$cols), selected = rv$analyte_map[[a]] %||% "")
    }))
  })
  observeEvent(input$auto_map, {
    req(rv$cols)
    for (k in names(analyte_syn)) {
      syns <- tolower(analyte_syn[[k]])
      hit <- rv$cols[tolower(rv$cols) %in% syns]
      if (length(hit)) updateSelectInput(session, paste0("map_", k), selected = hit[[1]])
    }
  })
  observe({
    req(rv$cols)
    ids <- names(reactiveValuesToList(input)); ids <- ids[grepl("^map_", ids)]
    amap <- list()
    for (id in ids) { key <- sub("^map_", "", id); val <- input[[id]]; if (isTruthy(val)) amap[[key]] <- val }
    rv$analyte_map <- amap
  })
  
  build_lb <- function() {
    req(rv$data, input$studyid, rv$animal_col, length(rv$analyte_map) > 0)
    df <- rv$data
    STUDYID <- input$studyid
    USUBJID_prefix <- if (isTruthy(input$usubid_col_prefix)) input$usubid_col_prefix else STUDYID
    ua_map <- c("-"=0, "NEG"=0, "+/-"=1, "1+"=1, "2+"=2, "3+"=3, "4+"=4)
    res <- list(); seq_counter <- 0L
    LBDY_fun <- function(date_value) {
      if (is.null(rv$date_col) || !nzchar(rv$date_col) || is.null(input$rfstdtc) || is.na(input$rfstdtc)) return(NA_integer_)
      d <- from_excel_date(date_value); if (is.na(d)) return(NA_integer_)
      as.integer(d - as.Date(input$rfstdtc)) + 1L
    }
    for (analyte in names(rv$analyte_map)) {
      src_col <- rv$analyte_map[[analyte]]; if (!nzchar(src_col) || !src_col %in% names(df)) next
      vals <- df[[src_col]]
      for (i in seq_len(nrow(df))) {
        seq_counter <- seq_counter + 1L
        animal <- df[[rv$animal_col]][i]
        usubjid <- sprintf("%s-%s", USUBJID_prefix, animal)
        orres <- as.character(vals[i]); orresu <- NA_character_; stresc <- orres
        if (grepl("^U_", analyte) && isTRUE(input$map_urine_semiquant)) {
          k <- toupper(trimws(orres)); k <- sub("^\\+$","1+",k)
          if (k %in% names(ua_map)) stresn <- ua_map[[k]] else stresn <- suppressWarnings(as.numeric(orres))
        } else {
          stresn <- suppressWarnings(as.numeric(orres))
        }
        lbdy <- if (!is.null(rv$date_col) && nzchar(rv$date_col)) LBDY_fun(df[[rv$date_col]][i]) else NA_integer_
        
        res[[length(res)+1L]] <- data.frame(
          STUDYID=STUDYID, DOMAIN="LB", USUBJID=usubjid, LBSEQ=seq_counter,
          LBTESTCD=analyte, LBTEST=analyte, LBCAT=input$category,
          LBORRES=orres, LBORRESU=orresu, LBSTRESC=stresc,
          LBSTRESN=ifelse(is.na(stresn), NA, as.numeric(stresn)),
          LBSTRESU=orresu, LBSPEC=input$specimen, LBBLFL=NA_character_,
          LBDY=lbdy, LBTPT=if (!is.null(rv$timepoint_col) && nzchar(rv$timepoint_col)) as.character(df[[rv$timepoint_col]][i]) else NA,
          LBTPTNUM=NA, stringsAsFactors = FALSE
        )
      }
    }
    if (!length(res)) return(NULL)
    lb <- do.call(rbind, res)
    lb <- lb[, c("STUDYID","DOMAIN","USUBJID","LBSEQ","LBTESTCD","LBTEST","LBCAT",
                 "LBORRES","LBORRESU","LBSTRESC","LBSTRESN","LBSTRESU","LBSPEC",
                 "LBBLFL","LBDY","LBTPT","LBTPTNUM"), drop = FALSE]
    lb
  }
  build_dm <- function() {
    req(input$studyid, input$rfstdtc, input$rfendtc)
    STUDYID <- input$studyid
    USUBID_col <- rv$animal_col
    animals <- if (!is.null(rv$data) && !is.null(USUBID_col) && nzchar(USUBID_col)) unique(as.character(na.omit(rv$data[[USUBID_col]]))) else character()
    USUBJID <- if (length(animals)) sprintf("%s-%s", if (isTruthy(input$usubid_col_prefix)) input$usubid_col_prefix else STUDYID, animals) else character()
    pick <- function(col, default) {
      if (!is.null(col) && nzchar(col) && !is.null(rv$data)) sapply(animals, function(a) {
        v <- rv$data[rv$data[[USUBID_col]]==a, col][1]; if (is.na(v) || !nzchar(v)) default else v
      }) else default %||% ""
    }
    data.frame(STUDYID=STUDYID, DOMAIN="DM", USUBJID=USUBJID, SUBJID=animals,
               RFSTDTC=as.character(as.Date(input$rfstdtc)),
               RFENDTC=as.character(as.Date(input$rfendtc)),
               SEX=pick(rv$sex_col, input$dm_sex_default),
               ARMCD=pick(rv$armcd_col, input$dm_armcd_default),
               SETCD=pick(rv$setcd_col, input$dm_setcd_default),
               stringsAsFactors = FALSE)
  }
  build_ts <- function() {
    req(input$studyid)
    data.frame(
      STUDYID=input$studyid, DOMAIN="TS", TSSEQ=1,
      TSPARMCD=c("SPECIES","STRAIN","ROUTE","SNDIGVER","SNDCTVER","STITLE","STSTDTC","STENDTC"),
      TSPARM=c("Species","Strain/Substrain","Route of Administration","SEND Implementation Guide Version",
               "SEND Controlled Terminology Version","Study Title","Study Start Date","Study End Date"),
      TSVAL=c(input$ts_species, input$ts_strain, input$ts_route, input$ts_sndigver, input$ts_sndctver, input$ts_title,
              as.character(as.Date(input$rfstdtc)), as.character(as.Date(input$rfendtc))),
      stringsAsFactors = FALSE
    )
  }
  build_ta <- function() {
    if (!isTruthy(input$ta_armcd) && !isTruthy(input$ta_arm))
      return(data.frame(STUDYID=input$studyid, DOMAIN="TA", ARMCD=character(), ARM=character(),
                        TAETORD=integer(), ETCD=character(), EPOCH=character()))
    data.frame(STUDYID=input$studyid, DOMAIN="TA", ARMCD=input$ta_armcd, ARM=input$ta_arm,
               TAETORD=1L, ETCD="TREATMENT", EPOCH="TREATMENT", stringsAsFactors = FALSE)
  }
  build_tx <- function() {
    if (!isTruthy(input$tx_setcd) && !isTruthy(input$tx_set))
      return(data.frame(STUDYID=input$studyid, DOMAIN="TX", SETCD=character(), SET=character(), TXSEQ=integer(),
                        TXPARMCD=character(), TXPARM=character(), TXVAL=character(), ARMCD=character(),
                        `Arm Code`=character(), SPGRPCD=character(), GRPLBL=character(), TRTDOS=character(),
                        TRTDOSU=character(), check.names = FALSE))
    data.frame(STUDYID=input$studyid, DOMAIN="TX", SETCD=input$tx_setcd, SET=input$tx_set,
               TXSEQ=1L, TXPARMCD="TRTDOS", TXPARM="Dose Level", TXVAL=input$tx_trtdos,
               ARMCD=input$ta_armcd, `Arm Code`=input$ta_armcd,
               SPGRPCD=NA, GRPLBL=NA, TRTDOS=input$tx_trtdos, TRTDOSU=input$tx_trtdosu,
               check.names = FALSE, stringsAsFactors = FALSE)
  }
  
  output$checks <- renderPrint({
    msgs <- c()
    if (!isTruthy(input$studyid)) msgs <- c(msgs, "[DM] STUDYID is required")
    if (is.null(input$rfstdtc) || is.na(input$rfstdtc)) msgs <- c(msgs, "[DM] RFSTDTC is required")
    if (is.null(input$rfendtc) || is.na(input$rfendtc)) msgs <- c(msgs, "[DM] RFENDTC is required")
    if (is.null(rv$data)) msgs <- c(msgs, "[LB] No data parsed yet")
    if (!isTruthy(rv$animal_col)) msgs <- c(msgs, "[LB] Animal ID column not selected")
    if (!length(rv$analyte_map)) msgs <- c(msgs, "[LB] No analytes mapped")
    if (!length(msgs)) {
      lb <- build_lb()
      if (!is.null(lb)) {
        if (any(!is.na(lb$LBSTRESN) & !is.numeric(lb$LBSTRESN))) msgs <- c(msgs, "[LB] STRESN must be numeric")
        if (any(!is.na(lb$LBDY) & (lb$LBDY != floor(lb$LBDY)))) msgs <- c(msgs, "[LB] --DY must be integer")
      }
    }
    if (!length(msgs)) {
      cat("All good ✔\n")
      if (!is.null(lb)) cat(sprintf("LB rows: %d\n", nrow(lb)))
    } else cat(paste(msgs, collapse = "\n"))
  })
  
  # ---- Browser export (single path used for Shinylive & classic Shiny) ----
  observeEvent(input$export_browser, {
    req(input$studyid)
    lb <- build_lb(); dm <- build_dm(); ts <- build_ts(); ta <- build_ta(); tx <- build_tx()
    if (is.null(lb) || !nrow(lb)) { showNotification("No LB rows to export. Check mappings.", type="error"); return() }
    payload <- list(
      filename = paste0("SEND_abbrev_", gsub("[^0-9A-Za-z_-]+","_", input$studyid), ".xlsx"),
      sheets   = list(LB=lb, DM=dm, TS=ts, TA=ta, TX=tx)
    )
    session$sendCustomMessage("download_xlsx",
                              jsonlite::toJSON(payload, dataframe="rows", auto_unbox=TRUE, na="null"))
  })
}

shinyApp(ui, server)