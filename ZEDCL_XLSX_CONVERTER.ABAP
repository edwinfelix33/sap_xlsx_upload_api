class ZEDCL_XLSX_CONVERTER definition
  public
  final
  create public .

public section.

  types:
    BEGIN OF GTY_SHEET_META_DATA,
    index type i,
    ref_id type string,
    name type string,
    end of GTY_SHEET_META_DATA .
  types:
    gtt_sheet_meta_data type table of gty_sheet_meta_data with EMPTY KEY .
  types:
    BEGIN OF gty_excel_columns,
         number TYPE i,
         name   TYPE string,
       END OF gty_excel_columns .
  types:
    gtt_excel_columns TYPE TABLE OF gty_excel_columns with EMPTY KEY .

  constants:
    BEGIN OF mc_xlsx_node ,
        string_item  TYPE string VALUE 'si',
        excel_row    TYPE string VALUE 'row',
        excel_column TYPE string VALUE 'c',
        excel_sheet  TYPE string VALUE 'sheet',
      END OF mc_xlsx_node .
  constants:
    BEGIN OF mc_xlsx_node_attribute ,
        cell_type TYPE string VALUE 't',
        cell_ref TYPE string VALUE 'r',
        row_no   TYPE string VALUE 'r',
        sheet_id TYPE string VALUE 'id',
        sheet_name TYPE string VALUE 'name',
      END OF mc_xlsx_node_attribute .
  constants:
    BEGIN OF mc_xlsx_attr_val,
                 string_type  TYPE string VALUE 's',
                 date_type    TYPE string VALUE 'd',
                 formula_type TYPE string VALUE 'f',
               END OF mc_xlsx_attr_val .
  constants:
    BEGIN OF MC_ABAP_TYPE,
    date_type type char1 VALUE 'D',
    time_type type char1 value 'T',
    end of MC_ABAP_TYPE .
  constants MC_LINE_BREAK type STRING value '<br/>' ##NO_TEXT.

  class-methods CONVERT_DATE_TO_INTERNAL
    importing
      !IV_DATE type STRING
    returning
      value(RV_DATE) type DATS .
  class-methods CONVERT_TEXT_DATE_TO_SAP
    importing
      !IV_DATE type STRING
    returning
      value(RV_DATE) type DATS .
  class-methods CONVERT_TEXT_TIME_TO_SAP
    importing
      !IV_TIME type STRING
    returning
      value(RV_TIME) type SYTIME .
  class-methods CONVERT_TIME_TO_INTERNAL
    importing
      !IV_TIME type STRING
    returning
      value(RV_TIME) type SYTIME .
  class-methods CONVERT_XLSX_DATE_TO_SAP
    importing
      !IV_DATE type STRING
    returning
      value(RV_DATE) type DATS .
  class-methods CONVERT_XLSX_TIME_TO_SAP
    importing
      !IV_TIME type STRING
    returning
      value(RV_TIME) type SYTIME .
  methods CONSTRUCTOR
    importing
      !IV_FILEPATH type STRING
      !IV_APPSERVER type BOOLE_D default ABAP_FALSE
      !IV_MAX_COLUMNS type I default 100
    exceptions
      GUI_UPLOAD_FAILED
      BIN_TO_XSTRING_FAILED .
  methods GET_DATA_AS_ITAB
    importing
      !IV_SHEET_INDEX type I optional
      !IV_SHEET_NAME type STRING optional
      !IV_SKIP_ROWS type I default 1
      !IV_CONV_ACTIVE type BOOLE_D default ABAP_TRUE
    exporting
      value(ET_DATA) type STANDARD TABLE
    exceptions
      INCOMPLETE_PARAMETERS .
  methods GET_MESSAGES
    returning
      value(RV_MESSAGES) type TT_RETURN .
  methods GET_TAB_SEPARATED_DATA
    importing
      !IV_SHEET_INDEX type I optional
      !IV_SHEET_NAME type STRING optional
    returning
      value(RV_DATA) type STRING
    exceptions
      INCOMPLETE_PARAMETERS .
  methods FREE .
protected section.
private section.

  data MV_XLSX_DATA type XSTRING .
  data MT_MESSAGES type TT_RETURN .
  data MV_DUMMY type CHAR1 .
  data MT_SHARED_STRINGS type ESH_T_STRING .
  data MT_EXCEL_COLUMNS type GTT_EXCEL_COLUMNS .
  data MT_SHEET_META_DATA type GTT_SHEET_META_DATA .

  methods FILL_EMPTY_ROWS
    importing
      !IV_PREV_ROW type I
      !IV_CURR_ROW type I
    changing
      !CV_DATA type STRING .
  methods ADD_MESSAGE .
  methods DOMAIN_CHECK .
  methods FILL_EMPTY_COLUMNS
    importing
      !IV_NEW_ROW type BOOLE_D
      !IV_PREV_COL type STRING
      !IV_CURR_COL type STRING
      !IV_SEPARATOR type ABAP_CHAR1
    changing
      !CV_DATA type STRING .
  methods GENERATE_EXCEL_COLUMN_NAMES
    importing
      !IV_MAX_COLUMNS type I
    returning
      value(RT_EXCEL_COLUMNS) type GTT_EXCEL_COLUMNS .
  methods GET_NEXT_EXCEL_COLUMN
    importing
      !IV_COLUMN type STRING
    returning
      value(RV_COLUMN) type STRING .
  methods GET_SHARED_STRINGS
    returning
      value(RT_SHARED_STRINGS) type ESH_T_STRING .
  methods GET_SHEET_DATA
    importing
      !IV_SHEET_REFID type STRING
      !IV_SEPARATOR type ABAP_CHAR1 default CL_ABAP_CHAR_UTILITIES=>HORIZONTAL_TAB
    returning
      value(RV_DATA) type STRING
    exceptions
      WORKSHEET_NOT_FOUND .
  methods GET_SHEET_META_DATA
    returning
      value(RT_SHEET_META_DATA) type GTT_SHEET_META_DATA .
  methods GET_XML_DOC_REF
    importing
      !IV_XML type STRING
    returning
      value(RO_IXML_DOCUMENT) type ref to IF_IXML_DOCUMENT .
ENDCLASS.



CLASS ZEDCL_XLSX_CONVERTER IMPLEMENTATION.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZEDCL_XLSX_CONVERTER->ADD_MESSAGE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD add_message.
    DATA: lw_message TYPE bapiret2.

    CHECK sy-msgty IS NOT INITIAL.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4 INTO lw_message-message.
    lw_message-id         = sy-msgid.
    lw_message-type       = sy-msgty.
    lw_message-number     = sy-msgno.
    lw_message-message_v1 = sy-msgv1.
    lw_message-message_v2 = sy-msgv2.
    lw_message-message_v3 = sy-msgv3.
    lw_message-message_v4 = sy-msgv4.
    APPEND lw_message TO mt_messages.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZEDCL_XLSX_CONVERTER->CONSTRUCTOR
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_FILEPATH                    TYPE        STRING
* | [--->] IV_APPSERVER                   TYPE        BOOLE_D (default =ABAP_FALSE)
* | [--->] IV_MAX_COLUMNS                 TYPE        I (default =100)
* | [EXC!] GUI_UPLOAD_FAILED
* | [EXC!] BIN_TO_XSTRING_FAILED
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD constructor.
    DATA: lw_data      TYPE solix,
          lt_data      TYPE solix_tab,
          lv_line_size TYPE i,
          lv_file_size TYPE i.

    IF iv_appserver EQ abap_true.
      TRY.
          OPEN DATASET iv_filepath FOR INPUT IN BINARY MODE.
          IF sy-subrc EQ 0.
            DO.
              READ DATASET iv_filepath INTO lw_data ACTUAL LENGTH lv_line_size.
              IF sy-subrc = 0.
                APPEND lw_data TO lt_data.
                ADD lv_line_size TO lv_file_size.
              ELSE.
                APPEND lw_data TO lt_data.
                ADD lv_line_size TO lv_file_size.
                EXIT.
              ENDIF.
            ENDDO.
            CLOSE DATASET iv_filepath.
          ENDIF.
        CATCH cx_root.
          "File open error
          RETURN.
      ENDTRY.
    ELSE.
      cl_gui_frontend_services=>gui_upload(
        EXPORTING
          filetype                = 'BIN'
          filename                = iv_filepath
      IMPORTING
        filelength = lv_file_size
        CHANGING
          data_tab                = lt_data
        EXCEPTIONS
          file_open_error         = 1
          file_read_error         = 2
          no_batch                = 3
          gui_refuse_filetransfer = 4
          invalid_type            = 5
          no_authority            = 6
          unknown_error           = 7
          bad_data_format         = 8
          header_not_allowed      = 9
          separator_not_allowed   = 10
          header_too_long         = 11
          unknown_dp_error        = 12
          access_denied           = 13
          dp_out_of_memory        = 14
          disk_full               = 15
          dp_timeout              = 16
          not_supported_by_gui    = 17
          error_no_gui            = 18
          OTHERS                  = 19
             ).
      IF sy-subrc <> 0.
        add_message( ).
        RETURN.
      ENDIF.
    ENDIF.
    CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
      EXPORTING
        input_length = lv_file_size
      IMPORTING
        buffer       = mv_xlsx_data
      TABLES
        binary_tab   = lt_data
      EXCEPTIONS
        failed       = 1
        OTHERS       = 2.
    IF sy-subrc <> 0.
      add_message( ).
      RETURN.
    ENDIF.
    mt_shared_strings   = me->get_shared_strings( ).
    mt_sheet_meta_data  = me->get_sheet_meta_data( ).
    mt_excel_columns    = me->generate_excel_column_names( iv_max_columns ).
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Static Public Method ZEDCL_XLSX_CONVERTER=>CONVERT_DATE_TO_INTERNAL
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_DATE                        TYPE        STRING
* | [<-()] RV_DATE                        TYPE        DATS
* +--------------------------------------------------------------------------------------</SIGNATURE>
method CONVERT_DATE_TO_INTERNAL.
  IF iv_date CO '0123456789'.
    rv_date = ZEDCL_XLSX_CONVERTER=>convert_xlsx_date_to_sap( iv_date ).
  ELSE.
    rv_date = ZEDCL_XLSX_CONVERTER=>convert_text_date_to_sap( iv_date ).
  ENDIF.
endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Static Public Method ZEDCL_XLSX_CONVERTER=>CONVERT_TEXT_DATE_TO_SAP
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_DATE                        TYPE        STRING
* | [<-()] RV_DATE                        TYPE        DATS
* +--------------------------------------------------------------------------------------</SIGNATURE>
METHOD convert_text_date_to_sap.
  IF iv_date IS NOT INITIAL.
    CALL FUNCTION 'CONVERT_DATE_TO_INTERNAL'
      EXPORTING
        date_external            = iv_date
      IMPORTING
        date_internal            = rv_date
      EXCEPTIONS
        date_external_is_invalid = 1
        OTHERS                   = 2.
    IF sy-subrc <> 0.
* Implement suitable error handling here
    ENDIF.
  ENDIF.

ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Static Public Method ZEDCL_XLSX_CONVERTER=>CONVERT_TEXT_TIME_TO_SAP
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_TIME                        TYPE        STRING
* | [<-()] RV_TIME                        TYPE        SYTIME
* +--------------------------------------------------------------------------------------</SIGNATURE>
METHOD convert_text_time_to_sap.
  CALL FUNCTION 'CONVERT_TIME_INPUT'
    EXPORTING
      input                     = iv_time
    IMPORTING
      output                    = rv_time
    EXCEPTIONS
      plausibility_check_failed = 1
      wrong_format_in_input     = 2
      OTHERS                    = 3.
  IF sy-subrc <> 0.

  ENDIF.

ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Static Public Method ZEDCL_XLSX_CONVERTER=>CONVERT_TIME_TO_INTERNAL
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_TIME                        TYPE        STRING
* | [<-()] RV_TIME                        TYPE        SYTIME
* +--------------------------------------------------------------------------------------</SIGNATURE>
METHOD convert_time_to_internal.
  IF iv_time CO '0123456789' AND iv_time CS '.'.
    rv_time = ZEDCL_XLSX_CONVERTER=>convert_xlsx_time_to_sap( iv_time ).
  ELSE.
    rv_time = ZEDCL_XLSX_CONVERTER=>convert_text_time_to_sap( iv_time ).
  ENDIF.
ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Static Public Method ZEDCL_XLSX_CONVERTER=>CONVERT_XLSX_DATE_TO_SAP
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_DATE                        TYPE        STRING
* | [<-()] RV_DATE                        TYPE        DATS
* +--------------------------------------------------------------------------------------</SIGNATURE>
METHOD convert_xlsx_date_to_sap.
  DATA: lv_offset TYPE i.

  CALL FUNCTION 'CONVERT_STRING_TO_INTEGER'
    EXPORTING
      p_string      = iv_date
    IMPORTING
      p_int         = lv_offset
    EXCEPTIONS
      overflow      = 1
      invalid_chars = 2
      OTHERS        = 3.
  IF sy-subrc <> 0 OR lv_offset EQ 0.
    RETURN.
  ENDIF.

  rv_date = '18991230'.
  ADD lv_offset TO rv_date.
*  rv_date = '19000101'.
*  ADD lv_offset TO rv_date.
*  SUBTRACT 2 FROM rv_date.
ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Static Public Method ZEDCL_XLSX_CONVERTER=>CONVERT_XLSX_TIME_TO_SAP
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_TIME                        TYPE        STRING
* | [<-()] RV_TIME                        TYPE        SYTIME
* +--------------------------------------------------------------------------------------</SIGNATURE>
method CONVERT_XLSX_TIME_TO_SAP.
  DATA: lv_float TYPE f,
        lv_hh TYPE numc2,
        lv_mm TYPE numc2,
        lv_ss TYPE numc2.

  lv_float = iv_time.
  lv_hh = floor( ( 86400 * lv_float ) / 3600 ).
  lv_mm = floor( ( ( 86400 * lv_float ) / 60 ) MOD 60 ).
  lv_ss = ( 86400 * lv_float ) MOD 100.
  CONCATENATE lv_hh lv_mm lv_ss INTO rv_time.
endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZEDCL_XLSX_CONVERTER->DOMAIN_CHECK
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method DOMAIN_CHECK.
  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZEDCL_XLSX_CONVERTER->FILL_EMPTY_COLUMNS
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_NEW_ROW                     TYPE        BOOLE_D
* | [--->] IV_PREV_COL                    TYPE        STRING
* | [--->] IV_CURR_COL                    TYPE        STRING
* | [--->] IV_SEPARATOR                   TYPE        ABAP_CHAR1
* | [<-->] CV_DATA                        TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
METHOD fill_empty_columns.
  DATA: lv_curr_col TYPE string,
        lv_prev_col TYPE string,
        lv_next_col TYPE string.
  lv_prev_col = iv_prev_col.
  lv_curr_col = iv_curr_col.

  SHIFT lv_prev_col RIGHT DELETING TRAILING '0123456789'.
  SHIFT lv_curr_col RIGHT DELETING TRAILING '0123456789'.

  CONDENSE lv_prev_col NO-GAPS.
  CONDENSE lv_curr_col NO-GAPS.

  IF iv_new_row EQ abap_true.
    IF lv_curr_col EQ 'A'.
      "New row first column OK
      RETURN.
    ELSE.
      lv_prev_col = 'A'.
    ENDIF.
  ELSE.
    lv_prev_col = me->get_next_excel_column( lv_prev_col ).
  ENDIF.

  WHILE lv_prev_col NE lv_curr_col.
    CONCATENATE cv_data iv_separator INTO cv_data.
    lv_prev_col = me->get_next_excel_column( lv_prev_col ).
  ENDWHILE.
ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZEDCL_XLSX_CONVERTER->FILL_EMPTY_ROWS
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_PREV_ROW                    TYPE        I
* | [--->] IV_CURR_ROW                    TYPE        I
* | [<-->] CV_DATA                        TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
METHOD fill_empty_rows.
  IF ( iv_curr_row - iv_prev_row ) GT 1.
    DO ( ( iv_curr_row - iv_prev_row ) - 1 ) TIMES.
      CONCATENATE cv_data cl_abap_char_utilities=>cr_lf INTO cv_data.
    ENDDO.
  ENDIF.
ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZEDCL_XLSX_CONVERTER->FREE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD free.
    FREE : mv_xlsx_data,
           mt_messages,
           mv_dummy,
           mt_shared_strings,
           mt_excel_columns,
           mt_sheet_meta_data.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZEDCL_XLSX_CONVERTER->GENERATE_EXCEL_COLUMN_NAMES
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_MAX_COLUMNS                 TYPE        I
* | [<-()] RT_EXCEL_COLUMNS               TYPE        GTT_EXCEL_COLUMNS
* +--------------------------------------------------------------------------------------</SIGNATURE>
METHOD generate_excel_column_names.
  DATA: lv_offset         TYPE i,
        lv_next_prefix    TYPE i,
        lv_countr         TYPE i VALUE 1,
        lv_prefix         TYPE string,
        lv_column         TYPE string,
        lv_max_col        TYPE i,
        ls_excel_columns  TYPE gty_excel_columns.

  lv_max_col = ceil( iv_max_columns / 26 ).
  DO lv_max_col TIMES.
    IF rt_excel_columns IS INITIAL.
      lv_prefix = ''.
      lv_next_prefix = 1.
    ELSE.
      READ TABLE rt_excel_columns INTO ls_excel_columns INDEX lv_next_prefix.
      lv_prefix = ls_excel_columns-name.
      ADD 1 TO lv_next_prefix.
    ENDIF.

    lv_offset = 0.
    DO 26 TIMES.
      CONCATENATE lv_prefix sy-abcde+lv_offset(1) INTO lv_column.
      APPEND VALUE #( number = lv_countr name = lv_column ) TO rt_excel_columns.
      IF lv_countr EQ iv_max_columns.
        EXIT.
      ENDIF.
      ADD 1 TO lv_offset.
      ADD 1 TO lv_countr.
    ENDDO.
  ENDDO.
ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZEDCL_XLSX_CONVERTER->GET_DATA_AS_ITAB
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_SHEET_INDEX                 TYPE        I(optional)
* | [--->] IV_SHEET_NAME                  TYPE        STRING(optional)
* | [--->] IV_SKIP_ROWS                   TYPE        I (default =1)
* | [--->] IV_CONV_ACTIVE                 TYPE        BOOLE_D (default =ABAP_TRUE)
* | [<---] ET_DATA                        TYPE        STANDARD TABLE
* | [EXC!] INCOMPLETE_PARAMETERS
* +--------------------------------------------------------------------------------------</SIGNATURE>
METHOD get_data_as_itab.
  DATA: lv_data       TYPE string,
        lv_cell       TYPE string,
        lt_data       TYPE TABLE OF string,
        lv_type       TYPE char1,
        lv_comp_count TYPE i,
        lv_row        TYPE i,
        lv_iterator   TYPE i.

  IF iv_sheet_index IS NOT SUPPLIED AND iv_sheet_name IS NOT SUPPLIED.
    RAISE incomplete_parameters.
  ENDIF.
  IF iv_sheet_name IS SUPPLIED.
    lv_data = me->get_tab_separated_data( iv_sheet_name = iv_sheet_name ).
  ELSE.
    lv_data = me->get_tab_separated_data( iv_sheet_index = iv_sheet_index ).
  ENDIF.
  SPLIT lv_data AT cl_abap_char_utilities=>cr_lf INTO TABLE lt_data.
  LOOP AT lt_data INTO lv_data FROM ( iv_skip_rows + 1 ).
    REPLACE ALL OCCURRENCES OF mc_line_break IN lv_data WITH cl_abap_char_utilities=>cr_lf.
    APPEND INITIAL LINE TO et_data ASSIGNING FIELD-SYMBOL(<lw_data>).
    DESCRIBE FIELD <lw_data> TYPE lv_type COMPONENTS lv_comp_count.
    lv_iterator = 1.
    DO ( lv_comp_count ) TIMES.
      ASSIGN COMPONENT lv_iterator OF STRUCTURE <lw_data> TO FIELD-SYMBOL(<lv_comp>).
      SPLIT lv_data AT cl_abap_char_utilities=>horizontal_tab INTO lv_cell lv_data.
      DESCRIBE FIELD <lv_comp> TYPE lv_type.
      TRY.
          IF iv_conv_active EQ abap_true.
            CASE  lv_type.
              WHEN mc_abap_type-date_type.
                <lv_comp> = ZEDCL_XLSX_CONVERTER=>convert_date_to_internal( iv_date = lv_cell ).
              WHEN mc_abap_type-time_type.
                <lv_comp> = ZEDCL_XLSX_CONVERTER=>convert_time_to_internal( iv_time = lv_cell ).
              WHEN OTHERS.
                <lv_comp> = lv_cell.
            ENDCASE.
          ELSE.
            <lv_comp> = lv_cell.
          ENDIF.
        CATCH cx_root.
          DESCRIBE TABLE et_data LINES lv_row.
          MESSAGE e003 WITH lv_row lv_iterator INTO mv_dummy.
          me->add_message( ).
      ENDTRY.
      ADD 1 TO lv_iterator.
    ENDDO.
  ENDLOOP.
ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZEDCL_XLSX_CONVERTER->GET_MESSAGES
* +-------------------------------------------------------------------------------------------------+
* | [<-()] RV_MESSAGES                    TYPE        TT_RETURN
* +--------------------------------------------------------------------------------------</SIGNATURE>
method GET_MESSAGES.
  rv_messages = me->mt_messages.
endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZEDCL_XLSX_CONVERTER->GET_NEXT_EXCEL_COLUMN
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_COLUMN                      TYPE        STRING
* | [<-()] RV_COLUMN                      TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
METHOD get_next_excel_column.
  DATA: lw_excel_column TYPE gty_excel_columns.

  READ TABLE me->mt_excel_columns INTO lw_excel_column with key name = iv_column.
  add 1 to lw_excel_column-number.
  READ TABLE me->mt_excel_columns INTO lw_excel_column INDEX lw_excel_column-number.
  rv_column = lw_excel_column-name.
ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZEDCL_XLSX_CONVERTER->GET_SHARED_STRINGS
* +-------------------------------------------------------------------------------------------------+
* | [<-()] RT_SHARED_STRINGS              TYPE        ESH_T_STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
METHOD get_shared_strings.
  DATA lo_xlsx_document             TYPE REF TO cl_xlsx_document.
  DATA lo_xlsx_workbook             TYPE REF TO cl_xlsx_workbookpart.
  DATA lo_xlsx_sharedstrings        TYPE REF TO cl_xlsx_sharedstringspart.
  DATA lo_hex_conv                  TYPE REF TO cl_abap_conv_in_ce.
  DATA lo_ixml_document             TYPE REF TO if_ixml_document.
  DATA lo_ixml_node_iterator        TYPE REF TO if_ixml_node_iterator.
  DATA lo_ixml_node_filter          TYPE REF TO if_ixml_node_filter.
  DATA lo_ixml_node                 TYPE REF TO if_ixml_node.

  DATA lv_hex_xml                   TYPE xstring.
  DATA lv_asc_xml                   TYPE string.

  TRY.
      "Get XLSX Document Reference
      lo_xlsx_document = cl_xlsx_document=>load_document( iv_data = me->mv_xlsx_data ).

      "Get XLSX Workbook Reference
      lo_xlsx_workbook = lo_xlsx_document->get_workbookpart( ).

      "Get Shared String Part
      "Info: Strings used in excel workbook are placed in a common file
      "      called shared_string.xml
      lo_xlsx_sharedstrings = lo_xlsx_workbook->get_sharedstringspart( ).

      "Get XML content of Shared String Part
      lv_hex_xml = lo_xlsx_sharedstrings->get_data( ).
      lo_hex_conv = cl_abap_conv_in_ce=>create( input = lv_hex_xml ).
      lo_hex_conv->read( IMPORTING data = lv_asc_xml ).

      "Validate XML & get ABAP XML Document Reference
      lo_ixml_document = me->get_xml_doc_ref( lv_asc_xml ).

      "Prepare XML node filter
      lo_ixml_node_filter = lo_ixml_document->create_filter_name( name = mc_xlsx_node-string_item ).

      "Get XML tree iterator
      lo_ixml_node_iterator = lo_ixml_document->create_iterator_filtered( lo_ixml_node_filter ).

      "Get Shared Strings
      lo_ixml_node = lo_ixml_node_iterator->get_next( ).
      WHILE NOT lo_ixml_node IS INITIAL.
        APPEND lo_ixml_node->get_value( ) TO rt_shared_strings.
        lo_ixml_node = lo_ixml_node_iterator->get_next( ).
      ENDWHILE.
    CATCH cx_openxml_not_found.
      me->add_message( ).
    CATCH cx_openxml_format.
      me->add_message( ).
  ENDTRY.
ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZEDCL_XLSX_CONVERTER->GET_SHEET_DATA
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_SHEET_REFID                 TYPE        STRING
* | [--->] IV_SEPARATOR                   TYPE        ABAP_CHAR1 (default =CL_ABAP_CHAR_UTILITIES=>HORIZONTAL_TAB)
* | [<-()] RV_DATA                        TYPE        STRING
* | [EXC!] WORKSHEET_NOT_FOUND
* +--------------------------------------------------------------------------------------</SIGNATURE>
METHOD get_sheet_data.
  DATA lo_xlsx_document             TYPE REF TO cl_xlsx_document.
  DATA lo_xlsx_workbook             TYPE REF TO cl_xlsx_workbookpart.
  DATA lo_xlsx_worksheet_collection TYPE REF TO cl_openxml_partcollection.
  DATA lo_openxml_worksheet         TYPE REF TO cl_openxml_part.
  DATA lo_hex_conv                  TYPE REF TO cl_abap_conv_in_ce.
  DATA lo_ixml_document             TYPE REF TO if_ixml_document.
  DATA lo_ixml_node_iterator        TYPE REF TO if_ixml_node_iterator.
  DATA lo_ixml_row_filter           TYPE REF TO if_ixml_node_filter.
  DATA lo_ixml_col_filter           TYPE REF TO if_ixml_node_filter.
  DATA lo_ixml_node_filter          TYPE REF TO if_ixml_node_filter.
  DATA lo_ixml_node                 TYPE REF TO if_ixml_node.
  DATA lv_hex_xml                   TYPE xstring.
  DATA lv_asc_xml                   TYPE string.
  DATA lv_attr_val                  TYPE string.
  DATA lv_new_row                   TYPE boole_d.
  DATA lv_prev_row                  TYPE i.
  DATA lv_curr_row                  TYPE i.
  DATA lv_prev_col                  TYPE string.
  DATA lv_curr_col                  TYPE string.
  DATA lv_si_index                  TYPE i. "String Item index (used to read from Shared String)

  TRY.
      "Get XLSX Document Reference
      lo_xlsx_document              = cl_xlsx_document=>load_document( iv_data = me->mv_xlsx_data ).

      "Get XLSX Workbook Reference
      lo_xlsx_workbook              = lo_xlsx_document->get_workbookpart( ).

      "Get Worksheet Part
      lo_xlsx_worksheet_collection  = lo_xlsx_workbook->get_worksheetparts( ).
      lo_openxml_worksheet          = lo_xlsx_workbook->get_part_by_id( iv_sheet_refid  ).

      "Get XML content of Worksheet
      lv_hex_xml                    = lo_openxml_worksheet->get_data( ).
      lo_hex_conv                   = cl_abap_conv_in_ce=>create( input = lv_hex_xml ).
      lo_hex_conv->read( IMPORTING data = lv_asc_xml ).

      "Validate XML & get ABAP XML Document Reference
      lo_ixml_document              = me->get_xml_doc_ref( lv_asc_xml ).

      "Prepare XML Node filter
      lo_ixml_row_filter            = lo_ixml_document->create_filter_name( name = mc_xlsx_node-excel_row ). "'row'
      lo_ixml_col_filter            = lo_ixml_document->create_filter_name( name = mc_xlsx_node-excel_column ).
      lo_ixml_node_filter           = lo_ixml_document->create_filter_or(
                  filter1           = lo_ixml_row_filter
                  filter2           = lo_ixml_col_filter  ).

      "Get XML Node Iterator (filter based)
      lo_ixml_node_iterator         = lo_ixml_document->create_iterator_filtered( lo_ixml_node_filter ).

      "Initialize values
      lv_prev_col = 0.

      "Get Sheet Data
      lo_ixml_node                  = lo_ixml_node_iterator->get_next( ).
      WHILE NOT lo_ixml_node IS INITIAL.
        "Get node (row/column) attribute map
          DATA(lo_attr_map)         = lo_ixml_node->get_attributes( ).
        IF lo_ixml_node->get_name( ) EQ mc_xlsx_node-excel_column.
          "Check for empty columns and fill them with empty values
          DATA(lo_attr_cell_ref)    = lo_attr_map->get_named_item_ns( name =  mc_xlsx_node_attribute-cell_ref ).
          IF lo_attr_cell_ref IS BOUND.
            lv_curr_col = lo_attr_cell_ref->get_value( ).
          ENDIF.
          me->fill_empty_columns(
            EXPORTING
              iv_new_row   = lv_new_row
              iv_prev_col  = lv_prev_col
              iv_curr_col  = lv_curr_col
              iv_separator = iv_separator
            CHANGING
              cv_data      = rv_data  ).

          "Get cell value
          DATA(lv_cell_val)         = lo_ixml_node->get_value( ).
          DATA(lo_attr_cell_type)   = lo_attr_map->get_named_item_ns( name =  mc_xlsx_node_attribute-cell_type ).
          IF lo_attr_cell_type IS BOUND.
            IF lo_attr_cell_type->get_value( ) EQ mc_xlsx_attr_val-string_type.
              "If string type get value from shared strings
              lv_si_index           = lv_cell_val.
              ADD 1 TO lv_si_index.
              READ TABLE mt_shared_strings INDEX lv_si_index INTO lv_cell_val.
            ENDIF.
          ENDIF.

          "Replace carriage return & line feed with line breaks. This is necessary so that TSV file isn't misaligned.
          replace all OCCURRENCES OF cl_abap_char_utilities=>cr_lf IN lv_cell_val with mc_line_break.

          "Write cell value to data string
          IF lv_new_row EQ abap_true.
            CONCATENATE rv_data lv_cell_val INTO rv_data.
          ELSE.
            CONCATENATE rv_data lv_cell_val INTO rv_data SEPARATED BY iv_separator.
          ENDIF.
          lv_new_row = abap_false.
          lv_prev_col = lv_curr_col.
        ELSE. "row
          DATA(lo_attr_row_no)    = lo_attr_map->get_named_item_ns( name =  mc_xlsx_node_attribute-row_no ).
          "Check for empty rows and fill them with empty CR_LF
          IF lo_attr_row_no IS BOUND.
            lv_attr_val = lo_attr_row_no->get_value( ).
            lv_curr_row = lv_attr_val.
          ENDIF.
          me->fill_empty_rows(
            EXPORTING
              iv_prev_row = lv_prev_row
              iv_curr_row = lv_curr_row
            CHANGING
              cv_data = rv_data
          ).

          "Append CR_LF only after first row
          IF lv_prev_row GT 0.
            CONCATENATE rv_data cl_abap_char_utilities=>cr_lf INTO rv_data.
          ENDIF.
          lv_new_row = abap_true.

          "Update previous Row/Column reference values
          CLEAR lv_prev_col.
          lv_prev_row = lv_curr_row.
        ENDIF.
        lo_ixml_node = lo_ixml_node_iterator->get_next( ).
      ENDWHILE.
    CATCH cx_openxml_not_found.
      me->add_message( ).
    CATCH cx_openxml_format.
      me->add_message( ).
    CATCH cx_root.
      IF lo_openxml_worksheet IS NOT BOUND.
        RAISE worksheet_not_found.
      ENDIF.
  ENDTRY.
  IF rv_data IS NOT INITIAL.
    CONCATENATE rv_data cl_abap_char_utilities=>cr_lf INTO rv_data.
  ENDIF.
ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZEDCL_XLSX_CONVERTER->GET_SHEET_META_DATA
* +-------------------------------------------------------------------------------------------------+
* | [<-()] RT_SHEET_META_DATA             TYPE        GTT_SHEET_META_DATA
* +--------------------------------------------------------------------------------------</SIGNATURE>
METHOD get_sheet_meta_data.
  DATA lo_xlsx_document             TYPE REF TO cl_xlsx_document.
  DATA lo_xlsx_workbook             TYPE REF TO cl_xlsx_workbookpart.
  DATA lo_hex_conv                  TYPE REF TO cl_abap_conv_in_ce.
  DATA lo_ixml_document             TYPE REF TO if_ixml_document.
  DATA lo_ixml_node_iterator        TYPE REF TO if_ixml_node_iterator.
  DATA lo_ixml_node_filter          TYPE REF TO if_ixml_node_filter.
  DATA lo_ixml_node                 TYPE REF TO if_ixml_node.
  DATA lv_hex_xml                   TYPE xstring.
  DATA lv_asc_xml                   TYPE string.
  DATA lv_index                     TYPE i.

  TRY.
      "Get XLSX Document Reference
      lo_xlsx_document              = cl_xlsx_document=>load_document( iv_data = me->mv_xlsx_data ).

      "Get XLSX Workbook Reference
      lo_xlsx_workbook              = lo_xlsx_document->get_workbookpart( ).

      "Get XML content of Workbook
      lv_hex_xml                    = lo_xlsx_workbook->get_data( ).
      lo_hex_conv                   = cl_abap_conv_in_ce=>create( input = lv_hex_xml ).
      lo_hex_conv->read( IMPORTING data = lv_asc_xml ).

      "Validate XML & get ABAP XML Document Reference
      lo_ixml_document              = me->get_xml_doc_ref( lv_asc_xml ).

      "Prepare XML node filter
      lo_ixml_node_filter           = lo_ixml_document->create_filter_name( name = mc_xlsx_node-excel_sheet ).

      "Get XML node iterator
      lo_ixml_node_iterator         = lo_ixml_document->create_iterator_filtered( lo_ixml_node_filter ).

      "Get Sheet Meta Data
      lo_ixml_node                  = lo_ixml_node_iterator->get_next( ).
      lv_index = 0.
      WHILE NOT lo_ixml_node IS INITIAL.
        DATA(lo_attr_map)         = lo_ixml_node->get_attributes( ).
        DATA(lo_attr_iterator)    = lo_attr_map->create_iterator( ).

        DATA(lo_attribute)        = lo_attr_iterator->get_next( ).
        WHILE lo_attribute IS NOT INITIAL.
          CASE lo_attribute->get_name( ).
            WHEN mc_xlsx_node_attribute-sheet_id.
              DATA(lv_sheet_ref_id) = lo_attribute->get_value( ).
            WHEN mc_xlsx_node_attribute-sheet_name.
              DATA(lv_sheet_name)   = lo_attribute->get_value( ).
            WHEN OTHERS.
          ENDCASE.
          lo_attribute            = lo_attr_iterator->get_next( ).
        ENDWHILE.

        APPEND VALUE #( index = lv_index ref_id = lv_sheet_ref_id name = lv_sheet_name ) TO rt_sheet_meta_data.

        lo_ixml_node = lo_ixml_node_iterator->get_next( ).
        ADD 1 TO lv_index.
      ENDWHILE.
    CATCH cx_openxml_not_found.
      me->add_message( ).
    CATCH cx_openxml_format.
      me->add_message( ).
  ENDTRY.
ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZEDCL_XLSX_CONVERTER->GET_TAB_SEPARATED_DATA
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_SHEET_INDEX                 TYPE        I(optional)
* | [--->] IV_SHEET_NAME                  TYPE        STRING(optional)
* | [<-()] RV_DATA                        TYPE        STRING
* | [EXC!] INCOMPLETE_PARAMETERS
* +--------------------------------------------------------------------------------------</SIGNATURE>
METHOD get_tab_separated_data.
  DATA: ls_sheet_meta_data type gty_sheet_meta_data.

  IF iv_sheet_index IS NOT SUPPLIED AND iv_sheet_name IS NOT SUPPLIED.
    RAISE incomplete_parameters.
  ENDIF.
  IF iv_sheet_name IS SUPPLIED.
    READ TABLE me->mt_sheet_meta_data INTO ls_sheet_meta_data WITH KEY name = iv_sheet_name.
  ELSE.
    READ TABLE me->mt_sheet_meta_data INTO ls_sheet_meta_data WITH KEY index = iv_sheet_index.
  ENDIF.
  rv_data = me->get_sheet_data( iv_sheet_refid  = ls_sheet_meta_data-ref_id
                                iv_separator    = cl_abap_char_utilities=>horizontal_tab ).
ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZEDCL_XLSX_CONVERTER->GET_XML_DOC_REF
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_XML                         TYPE        STRING
* | [<-()] RO_IXML_DOCUMENT               TYPE REF TO IF_IXML_DOCUMENT
* +--------------------------------------------------------------------------------------</SIGNATURE>
METHOD get_xml_doc_ref.
  DATA : lo_ixml                TYPE REF TO if_ixml,
         lo_ixml_document       TYPE REF TO if_ixml_document,
         lo_ixml_stream_factory TYPE REF TO if_ixml_stream_factory,
         lo_ixml_istream        TYPE REF TO if_ixml_istream,
         lo_ixml_parser         TYPE REF TO if_ixml_parser,
         lo_ixml_parse_error    TYPE REF TO if_ixml_parse_error,
         lo_ixml_node           TYPE REF TO if_ixml_node,

         lv_index               TYPE i.

  "Create the Main Factory
  lo_ixml = cl_ixml=>create( ).

  "Create the Initial Document
  lo_ixml_document = lo_ixml->create_document( ).

  "Create a Stream Factory
  lo_ixml_stream_factory = lo_ixml->create_stream_factory( ).

  "Create an Input Stream
  lo_ixml_istream       = lo_ixml_stream_factory->create_istream_string( string = iv_xml ).

  "Create a Parser
  lo_ixml_parser = lo_ixml->create_parser(
      document       = lo_ixml_document
      istream        = lo_ixml_istream
      stream_factory = lo_ixml_stream_factory
  ).

  "Check for XML parsing errors
  IF lo_ixml_parser->parse( ) <> 0.
    lv_index = 0.
    WHILE lv_index < lo_ixml_parser->num_errors( ).
      lo_ixml_parse_error = lo_ixml_parser->get_error( index = lv_index ).
      "XML error at row &1 column &2: &3
      MESSAGE e002 WITH lo_ixml_parse_error->get_line( ) lo_ixml_parse_error->get_column( ) lo_ixml_parse_error->get_reason( ) INTO mv_dummy.
      me->add_message( ).
      ADD 1 TO lv_index.
    ENDWHILE.
  ENDIF.
  lo_ixml_istream->close( ).

  ro_ixml_document = lo_ixml_document.
ENDMETHOD.
ENDCLASS.
