*&---------------------------------------------------------------------*
*& Report ZRAPORT_WYDAJNOSCI
*&---------------------------------------------------------------------*
*& Raport wydajności dla wydziału produkcyjnego
*& - Wczytywanie danych z plików TSV
*& - Wyświetlanie w ALV
*& - Podgląd HTM
*& - Wysyłanie maili z HTM/XLSX
*&---------------------------------------------------------------------*
REPORT zraport_wydajnosci.

*----------------------------------------------------------------------*
* TABLES
*----------------------------------------------------------------------*
TABLES: sscrfields.

*----------------------------------------------------------------------*
* TYPE DEFINITIONS
*----------------------------------------------------------------------*
TYPES: BEGIN OF ty_wydajnosc,
         wydzial      TYPE c LENGTH 40,     " Wydział produkcyjny
         data         TYPE d,                " Data
         pracownik    TYPE c LENGTH 100,    " Pracownik
         produkt      TYPE c LENGTH 40,     " Produkt
         ilosc        TYPE i,                " Ilość wyprodukowana
         czas_pracy   TYPE p DECIMALS 2,    " Czas pracy (godziny)
         wydajnosc    TYPE p DECIMALS 2,    " Wydajność (szt/godz)
         cel          TYPE p DECIMALS 2,    " Cel wydajności
         procent_celu TYPE p DECIMALS 2,    " Procent realizacji celu
         uwagi        TYPE c LENGTH 255,    " Uwagi
       END OF ty_wydajnosc.

TYPES: tt_wydajnosc TYPE STANDARD TABLE OF ty_wydajnosc WITH DEFAULT KEY.

*----------------------------------------------------------------------*
* DATA DECLARATIONS
*----------------------------------------------------------------------*
DATA: gt_wydajnosc   TYPE tt_wydajnosc,
      go_alv         TYPE REF TO cl_salv_table,
      go_container   TYPE REF TO cl_gui_custom_container,
      gv_file_path   TYPE string,
      gv_wydzial     TYPE c LENGTH 40.

*----------------------------------------------------------------------*
* SELECTION SCREEN
*----------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-001.
  PARAMETERS: p_file   TYPE localfile OBLIGATORY,
              p_wydzl  TYPE c LENGTH 40 OBLIGATORY.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-002.
  PARAMETERS: p_htm    RADIOBUTTON GROUP rb1,
              p_xlsx   RADIOBUTTON GROUP rb1 DEFAULT 'X'.
SELECTION-SCREEN END OF BLOCK b2.

SELECTION-SCREEN FUNCTION KEY 1.
SELECTION-SCREEN FUNCTION KEY 2.
SELECTION-SCREEN FUNCTION KEY 3.

*----------------------------------------------------------------------*
* AT SELECTION-SCREEN
*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  PERFORM f4_help_file CHANGING p_file.

AT SELECTION-SCREEN.
  CASE sscrfields-ucomm.
    WHEN 'FC01'.
      " Podgląd HTM
      PERFORM show_htm_preview.
    WHEN 'FC02'.
      " Wyślij mail z HTM
      PERFORM send_email USING 'HTM'.
    WHEN 'FC03'.
      " Wyślij mail z XLSX
      PERFORM send_email USING 'XLSX'.
  ENDCASE.

*----------------------------------------------------------------------*
* INITIALIZATION
*----------------------------------------------------------------------*
INITIALIZATION.
  TEXT-001 = 'Parametry wejściowe'.
  TEXT-002 = 'Format eksportu'.

  sscrfields-functxt_01 = 'Podgląd HTM'.
  sscrfields-functxt_02 = 'Wyślij HTM'.
  sscrfields-functxt_03 = 'Wyślij XLSX'.

*----------------------------------------------------------------------*
* START-OF-SELECTION
*----------------------------------------------------------------------*
START-OF-SELECTION.
  PERFORM load_tsv_data.
  PERFORM calculate_efficiency.
  PERFORM display_alv.

*&---------------------------------------------------------------------*
*& Form F4_HELP_FILE
*&---------------------------------------------------------------------*
FORM f4_help_file CHANGING cv_file TYPE localfile.
  DATA: lt_file_table TYPE filetable,
        ls_file       TYPE file_table,
        lv_rc         TYPE i,
        lv_action     TYPE i.

  CALL METHOD cl_gui_frontend_services=>file_open_dialog
    EXPORTING
      window_title            = 'Wybierz plik TSV'
      file_filter             = 'TSV Files (*.tsv)|*.tsv|All Files (*.*)|*.*'
      initial_directory       = 'C:\'
    CHANGING
      file_table              = lt_file_table
      rc                      = lv_rc
      user_action             = lv_action
    EXCEPTIONS
      file_open_dialog_failed = 1
      cntl_error              = 2
      error_no_gui            = 3
      not_supported_by_gui    = 4
      OTHERS                  = 5.

  IF sy-subrc = 0 AND lv_action = cl_gui_frontend_services=>action_ok.
    READ TABLE lt_file_table INTO ls_file INDEX 1.
    IF sy-subrc = 0.
      cv_file = ls_file-filename.
    ENDIF.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form LOAD_TSV_DATA
*&---------------------------------------------------------------------*
FORM load_tsv_data.
  DATA: lt_raw_data TYPE TABLE OF string,
        lv_line     TYPE string,
        lt_fields   TYPE TABLE OF string,
        lv_field    TYPE string,
        ls_data     TYPE ty_wydajnosc,
        lv_first    TYPE abap_bool VALUE abap_true.

  " Wczytaj plik TSV
  CALL METHOD cl_gui_frontend_services=>gui_upload
    EXPORTING
      filename                = CONV #( p_file )
      filetype                = 'ASC'
      codepage                = '4110'
    CHANGING
      data_tab                = lt_raw_data
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
      OTHERS                  = 17.

  IF sy-subrc <> 0.
    MESSAGE e001(00) WITH 'Błąd wczytywania pliku TSV'.
    RETURN.
  ENDIF.

  " Parsuj dane TSV
  LOOP AT lt_raw_data INTO lv_line.
    " Pomiń nagłówek
    IF lv_first = abap_true.
      lv_first = abap_false.
      CONTINUE.
    ENDIF.

    " Pomiń puste linie
    IF lv_line IS INITIAL.
      CONTINUE.
    ENDIF.

    CLEAR: lt_fields, ls_data.
    SPLIT lv_line AT cl_abap_char_utilities=>horizontal_tab INTO TABLE lt_fields.

    " Mapuj pola
    READ TABLE lt_fields INTO lv_field INDEX 1.
    IF sy-subrc = 0. ls_data-wydzial = lv_field. ENDIF.

    READ TABLE lt_fields INTO lv_field INDEX 2.
    IF sy-subrc = 0. ls_data-data = lv_field. ENDIF.

    READ TABLE lt_fields INTO lv_field INDEX 3.
    IF sy-subrc = 0. ls_data-pracownik = lv_field. ENDIF.

    READ TABLE lt_fields INTO lv_field INDEX 4.
    IF sy-subrc = 0. ls_data-produkt = lv_field. ENDIF.

    READ TABLE lt_fields INTO lv_field INDEX 5.
    IF sy-subrc = 0. ls_data-ilosc = lv_field. ENDIF.

    READ TABLE lt_fields INTO lv_field INDEX 6.
    IF sy-subrc = 0. ls_data-czas_pracy = lv_field. ENDIF.

    READ TABLE lt_fields INTO lv_field INDEX 7.
    IF sy-subrc = 0. ls_data-cel = lv_field. ENDIF.

    READ TABLE lt_fields INTO lv_field INDEX 8.
    IF sy-subrc = 0. ls_data-uwagi = lv_field. ENDIF.

    " Filtruj według wydziału jeśli podano (bez rozróżniania wielkości liter)
    IF p_wydzl IS NOT INITIAL.
      DATA(lv_wydzial_upper) = to_upper( ls_data-wydzial ).
      DATA(lv_filter_upper) = to_upper( CONV string( p_wydzl ) ).
      IF lv_wydzial_upper CS lv_filter_upper.
        APPEND ls_data TO gt_wydajnosc.
      ENDIF.
    ELSE.
      APPEND ls_data TO gt_wydajnosc.
    ENDIF.
  ENDLOOP.

  IF gt_wydajnosc IS INITIAL.
    MESSAGE s002(00) WITH 'Brak danych dla wybranego wydziału' DISPLAY LIKE 'W'.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form CALCULATE_EFFICIENCY
*&---------------------------------------------------------------------*
FORM calculate_efficiency.
  FIELD-SYMBOLS: <ls_data> TYPE ty_wydajnosc.

  LOOP AT gt_wydajnosc ASSIGNING <ls_data>.
    " Oblicz wydajność (sztuk na godzinę)
    IF <ls_data>-czas_pracy > 0.
      <ls_data>-wydajnosc = <ls_data>-ilosc / <ls_data>-czas_pracy.
    ENDIF.

    " Oblicz procent realizacji celu
    IF <ls_data>-cel > 0.
      <ls_data>-procent_celu = ( <ls_data>-wydajnosc / <ls_data>-cel ) * 100.
    ENDIF.
  ENDLOOP.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form DISPLAY_ALV
*&---------------------------------------------------------------------*
FORM display_alv.
  DATA: lo_columns    TYPE REF TO cl_salv_columns_table,
        lo_column     TYPE REF TO cl_salv_column,
        lo_functions  TYPE REF TO cl_salv_functions,
        lo_display    TYPE REF TO cl_salv_display_settings,
        lo_events     TYPE REF TO cl_salv_events_table,
        lo_layout     TYPE REF TO cl_salv_layout,
        ls_layout_key TYPE salv_s_layout_key.

  TRY.
      " Utwórz instancję ALV
      cl_salv_table=>factory(
        IMPORTING
          r_salv_table = go_alv
        CHANGING
          t_table      = gt_wydajnosc ).

      " Włącz wszystkie funkcje standardowe
      lo_functions = go_alv->get_functions( ).
      lo_functions->set_all( abap_true ).

      " Dodaj własne funkcje
      lo_functions->add_function(
        name     = 'HTM_PREVIEW'
        icon     = '@3Y@'
        text     = 'Podgląd HTM'
        tooltip  = 'Podgląd raportu w formacie HTM'
        position = if_salv_c_function_position=>right_of_salv_functions ).

      lo_functions->add_function(
        name     = 'SEND_HTM'
        icon     = '@48@'
        text     = 'Wyślij HTM'
        tooltip  = 'Wyślij raport mailem w formacie HTM'
        position = if_salv_c_function_position=>right_of_salv_functions ).

      lo_functions->add_function(
        name     = 'SEND_XLSX'
        icon     = '@49@'
        text     = 'Wyślij XLSX'
        tooltip  = 'Wyślij raport mailem w formacie XLSX'
        position = if_salv_c_function_position=>right_of_salv_functions ).

      " Konfiguruj kolumny
      lo_columns = go_alv->get_columns( ).
      lo_columns->set_optimize( abap_true ).

      PERFORM set_column_texts USING lo_columns.

      " Ustaw tytuł
      lo_display = go_alv->get_display_settings( ).
      lo_display->set_list_header( |Raport wydajności - Wydział: { p_wydzl }| ).
      lo_display->set_striped_pattern( abap_true ).

      " Włącz layout
      lo_layout = go_alv->get_layout( ).
      ls_layout_key-report = sy-repid.
      lo_layout->set_key( ls_layout_key ).
      lo_layout->set_save_restriction( if_salv_c_layout=>restrict_none ).

      " Obsługa zdarzeń
      lo_events = go_alv->get_event( ).
      SET HANDLER lcl_event_handler=>on_user_command FOR lo_events.
      SET HANDLER lcl_event_handler=>on_double_click FOR lo_events.

      " Wyświetl ALV
      go_alv->display( ).

    CATCH cx_salv_msg INTO DATA(lx_msg).
      MESSAGE lx_msg TYPE 'E'.
    CATCH cx_salv_not_found INTO DATA(lx_not_found).
      MESSAGE lx_not_found TYPE 'E'.
    CATCH cx_salv_wrong_call INTO DATA(lx_wrong_call).
      MESSAGE lx_wrong_call TYPE 'E'.
    CATCH cx_salv_existing INTO DATA(lx_existing).
      MESSAGE lx_existing TYPE 'E'.
  ENDTRY.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form SET_COLUMN_TEXTS
*&---------------------------------------------------------------------*
FORM set_column_texts USING io_columns TYPE REF TO cl_salv_columns_table.
  DATA: lo_column TYPE REF TO cl_salv_column.

  TRY.
      lo_column = io_columns->get_column( 'WYDZIAL' ).
      lo_column->set_short_text( 'Wydział' ).
      lo_column->set_medium_text( 'Wydział' ).
      lo_column->set_long_text( 'Wydział produkcyjny' ).

      lo_column = io_columns->get_column( 'DATA' ).
      lo_column->set_short_text( 'Data' ).
      lo_column->set_medium_text( 'Data' ).
      lo_column->set_long_text( 'Data produkcji' ).

      lo_column = io_columns->get_column( 'PRACOWNIK' ).
      lo_column->set_short_text( 'Pracownik' ).
      lo_column->set_medium_text( 'Pracownik' ).
      lo_column->set_long_text( 'Imię i nazwisko pracownika' ).

      lo_column = io_columns->get_column( 'PRODUKT' ).
      lo_column->set_short_text( 'Produkt' ).
      lo_column->set_medium_text( 'Produkt' ).
      lo_column->set_long_text( 'Nazwa produktu' ).

      lo_column = io_columns->get_column( 'ILOSC' ).
      lo_column->set_short_text( 'Ilość' ).
      lo_column->set_medium_text( 'Ilość' ).
      lo_column->set_long_text( 'Ilość wyprodukowana' ).

      lo_column = io_columns->get_column( 'CZAS_PRACY' ).
      lo_column->set_short_text( 'Czas[h]' ).
      lo_column->set_medium_text( 'Czas pracy[h]' ).
      lo_column->set_long_text( 'Czas pracy w godzinach' ).

      lo_column = io_columns->get_column( 'WYDAJNOSC' ).
      lo_column->set_short_text( 'Wydajn.' ).
      lo_column->set_medium_text( 'Wydajność' ).
      lo_column->set_long_text( 'Wydajność (szt/godz)' ).

      lo_column = io_columns->get_column( 'CEL' ).
      lo_column->set_short_text( 'Cel' ).
      lo_column->set_medium_text( 'Cel wydajności' ).
      lo_column->set_long_text( 'Cel wydajności (szt/godz)' ).

      lo_column = io_columns->get_column( 'PROCENT_CELU' ).
      lo_column->set_short_text( '%Celu' ).
      lo_column->set_medium_text( '% Celu' ).
      lo_column->set_long_text( 'Procent realizacji celu' ).

      lo_column = io_columns->get_column( 'UWAGI' ).
      lo_column->set_short_text( 'Uwagi' ).
      lo_column->set_medium_text( 'Uwagi' ).
      lo_column->set_long_text( 'Uwagi dodatkowe' ).

    CATCH cx_salv_not_found.
      " Ignoruj błędy dla nieistniejących kolumn
  ENDTRY.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form SHOW_HTM_PREVIEW
*&---------------------------------------------------------------------*
FORM show_htm_preview.
  DATA: lv_html    TYPE string,
        lv_file    TYPE string,
        lt_html    TYPE TABLE OF string.

  " Wygeneruj HTML
  PERFORM generate_html CHANGING lv_html.

  " Pobierz katalog tymczasowy i zapisz plik
  DATA: lv_temp_dir TYPE string.

  CALL METHOD cl_gui_frontend_services=>get_temp_directory
    CHANGING
      temp_dir             = lv_temp_dir
    EXCEPTIONS
      cntl_error           = 1
      error_no_gui         = 2
      not_supported_by_gui = 3
      OTHERS               = 4.

  IF sy-subrc = 0.
    lv_file = lv_temp_dir && '\raport_wydajnosci.htm'.
  ELSE.
    lv_file = 'C:\temp\raport_wydajnosci.htm'.
  ENDIF.

  SPLIT lv_html AT cl_abap_char_utilities=>cr_lf INTO TABLE lt_html.

  CALL METHOD cl_gui_frontend_services=>gui_download
    EXPORTING
      filename              = lv_file
      filetype              = 'ASC'
      codepage              = '4110'
    CHANGING
      data_tab              = lt_html
    EXCEPTIONS
      file_write_error      = 1
      no_batch              = 2
      gui_refuse_filetransfer = 3
      invalid_type          = 4
      no_authority          = 5
      unknown_error         = 6
      header_not_allowed    = 7
      separator_not_allowed = 8
      filesize_not_allowed  = 9
      header_too_long       = 10
      dp_error_create       = 11
      dp_error_send         = 12
      dp_error_write        = 13
      unknown_dp_error      = 14
      access_denied         = 15
      dp_out_of_memory      = 16
      disk_full             = 17
      dp_timeout            = 18
      file_not_found        = 19
      dataprovider_exception = 20
      control_flush_error   = 21
      OTHERS                = 22.

  IF sy-subrc = 0.
    " Otwórz w przeglądarce
    CALL METHOD cl_gui_frontend_services=>execute
      EXPORTING
        document               = lv_file
        operation              = 'OPEN'
      EXCEPTIONS
        cntl_error             = 1
        error_no_gui           = 2
        bad_parameter          = 3
        file_not_found         = 4
        path_not_found         = 5
        file_extension_unknown = 6
        error_execute_failed   = 7
        synchronous_failed     = 8
        not_supported_by_gui   = 9
        OTHERS                 = 10.

    IF sy-subrc <> 0.
      MESSAGE s003(00) WITH 'Nie można otworzyć pliku HTM' DISPLAY LIKE 'E'.
    ENDIF.
  ELSE.
    MESSAGE s004(00) WITH 'Błąd zapisu pliku HTM' DISPLAY LIKE 'E'.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form GENERATE_HTML
*&---------------------------------------------------------------------*
FORM generate_html CHANGING cv_html TYPE string.
  DATA: lv_row  TYPE string,
        ls_data TYPE ty_wydajnosc.

  " Nagłówek HTML
  cv_html = |<!DOCTYPE html>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |<html lang="pl">{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |<head>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |<meta charset="UTF-8">{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |<title>Raport Wydajności - { p_wydzl }</title>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |<style>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |body \{ font-family: Arial, sans-serif; margin: 20px; \}{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |h1 \{ color: #333; \}{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |table \{ border-collapse: collapse; width: 100%; margin-top: 20px; \}{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |th, td \{ border: 1px solid #ddd; padding: 8px; text-align: left; \}{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |th \{ background-color: #4CAF50; color: white; \}{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |tr:nth-child(even) \{ background-color: #f2f2f2; \}{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |tr:hover \{ background-color: #ddd; \}{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |.high \{ color: green; font-weight: bold; \}{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |.low \{ color: red; font-weight: bold; \}{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |</style>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |</head>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |<body>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |<h1>Raport Wydajności - Wydział: { p_wydzl }</h1>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |<p>Data wygenerowania: { sy-datum } { sy-uzeit }</p>{ cl_abap_char_utilities=>cr_lf }|.

  " Tabela
  cv_html = cv_html && |<table>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |<thead>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |<tr>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |<th>Wydział</th>|.
  cv_html = cv_html && |<th>Data</th>|.
  cv_html = cv_html && |<th>Pracownik</th>|.
  cv_html = cv_html && |<th>Produkt</th>|.
  cv_html = cv_html && |<th>Ilość</th>|.
  cv_html = cv_html && |<th>Czas pracy [h]</th>|.
  cv_html = cv_html && |<th>Wydajność</th>|.
  cv_html = cv_html && |<th>Cel</th>|.
  cv_html = cv_html && |<th>% Celu</th>|.
  cv_html = cv_html && |<th>Uwagi</th>|.
  cv_html = cv_html && |</tr>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |</thead>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |<tbody>{ cl_abap_char_utilities=>cr_lf }|.

  " Wiersze danych
  LOOP AT gt_wydajnosc INTO ls_data.
    DATA(lv_class) = COND string(
      WHEN ls_data-procent_celu >= 100 THEN 'high'
      WHEN ls_data-procent_celu < 80 THEN 'low'
      ELSE '' ).

    cv_html = cv_html && |<tr>{ cl_abap_char_utilities=>cr_lf }|.
    cv_html = cv_html && |<td>{ ls_data-wydzial }</td>|.
    cv_html = cv_html && |<td>{ ls_data-data+6(2) }.{ ls_data-data+4(2) }.{ ls_data-data(4) }</td>|.
    cv_html = cv_html && |<td>{ ls_data-pracownik }</td>|.
    cv_html = cv_html && |<td>{ ls_data-produkt }</td>|.
    cv_html = cv_html && |<td>{ ls_data-ilosc }</td>|.
    cv_html = cv_html && |<td>{ ls_data-czas_pracy }</td>|.
    cv_html = cv_html && |<td>{ ls_data-wydajnosc }</td>|.
    cv_html = cv_html && |<td>{ ls_data-cel }</td>|.
    cv_html = cv_html && |<td class="{ lv_class }">{ ls_data-procent_celu }%</td>|.
    cv_html = cv_html && |<td>{ ls_data-uwagi }</td>|.
    cv_html = cv_html && |</tr>{ cl_abap_char_utilities=>cr_lf }|.
  ENDLOOP.

  cv_html = cv_html && |</tbody>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |</table>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |</body>{ cl_abap_char_utilities=>cr_lf }|.
  cv_html = cv_html && |</html>{ cl_abap_char_utilities=>cr_lf }|.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form SEND_EMAIL
*&---------------------------------------------------------------------*
FORM send_email USING iv_format TYPE string.
  DATA: lo_send_request  TYPE REF TO cl_bcs,
        lo_document      TYPE REF TO cl_document_bcs,
        lo_sender        TYPE REF TO cl_sapuser_bcs,
        lo_recipient     TYPE REF TO if_recipient_bcs,
        lt_content       TYPE soli_tab,
        lv_subject       TYPE so_obj_des,
        lv_html          TYPE string,
        lt_attachment    TYPE solix_tab,
        lv_att_name      TYPE sood-objnam,
        lv_att_type      TYPE sood-objtp,
        lv_size          TYPE sood-objlen,
        lv_email         TYPE adr6-smtp_addr.

  " Zapytaj o adres email
  CALL FUNCTION 'POPUP_TO_GET_VALUE'
    EXPORTING
      fieldname        = 'EMAIL'
      tabname          = 'ADR6'
      titel            = 'Podaj adres email odbiorcy'
      valuein          = ''
    IMPORTING
      valueout         = lv_email
    EXCEPTIONS
      fieldname_not_found = 1
      OTHERS           = 2.

  IF sy-subrc <> 0 OR lv_email IS INITIAL.
    MESSAGE s005(00) WITH 'Anulowano wysyłanie maila' DISPLAY LIKE 'W'.
    RETURN.
  ENDIF.

  TRY.
      " Utwórz request wysyłki
      lo_send_request = cl_bcs=>create_persistent( ).

      " Przygotuj treść maila
      APPEND 'Szanowni Państwo,' TO lt_content.
      APPEND '' TO lt_content.
      APPEND |W załączeniu przesyłam raport wydajności dla wydziału: { p_wydzl }| TO lt_content.
      APPEND '' TO lt_content.
      APPEND 'Z poważaniem,' TO lt_content.
      APPEND sy-uname TO lt_content.

      lv_subject = |Raport wydajności - { p_wydzl } - { sy-datum }|.

      " Utwórz dokument
      lo_document = cl_document_bcs=>create_document(
        i_type    = 'RAW'
        i_text    = lt_content
        i_subject = lv_subject ).

      " Dodaj załącznik
      lv_att_name = 'raport_wydajnosci'.
      IF iv_format = 'HTM'.
        PERFORM generate_html CHANGING lv_html.
        PERFORM convert_string_to_solix USING lv_html CHANGING lt_attachment lv_size.
        lv_att_type = 'HTM'.
      ELSE.
        PERFORM generate_xlsx CHANGING lt_attachment lv_size.
        lv_att_type = 'XLS'.
      ENDIF.

      lo_document->add_attachment(
        i_attachment_type    = lv_att_type
        i_attachment_subject = CONV #( lv_att_name )
        i_att_content_hex    = lt_attachment ).

      lo_send_request->set_document( lo_document ).

      " Ustaw nadawcę
      lo_sender = cl_sapuser_bcs=>create( sy-uname ).
      lo_send_request->set_sender( lo_sender ).

      " Dodaj odbiorcę
      lo_recipient = cl_cam_address_bcs=>create_internet_address( lv_email ).
      lo_send_request->add_recipient(
        i_recipient  = lo_recipient
        i_express    = abap_true ).

      " Wyślij natychmiast
      lo_send_request->set_send_immediately( abap_true ).

      " Wyślij
      lo_send_request->send( i_with_error_screen = abap_true ).

      COMMIT WORK.

      MESSAGE s006(00) WITH 'Mail został wysłany pomyślnie'.

    CATCH cx_bcs INTO DATA(lx_bcs).
      MESSAGE lx_bcs TYPE 'E'.
  ENDTRY.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form CONVERT_STRING_TO_SOLIX
*&---------------------------------------------------------------------*
FORM convert_string_to_solix
  USING    iv_string TYPE string
  CHANGING ct_solix  TYPE solix_tab
           cv_size   TYPE sood-objlen.

  DATA: lv_xstring TYPE xstring.

  " Konwertuj string na xstring
  CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
    EXPORTING
      text   = iv_string
    IMPORTING
      buffer = lv_xstring
    EXCEPTIONS
      failed = 1
      OTHERS = 2.

  IF sy-subrc = 0.
    " Konwertuj xstring na solix
    ct_solix = cl_bcs_convert=>xstring_to_solix( iv_xstring = lv_xstring ).
    cv_size = xstrlen( lv_xstring ).
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form GENERATE_XLSX
*&---------------------------------------------------------------------*
FORM generate_xlsx
  CHANGING ct_data TYPE solix_tab
           cv_size TYPE sood-objlen.

  DATA: lo_excel     TYPE REF TO cl_salv_table,
        lv_xstring   TYPE xstring.

  TRY.
      " Utwórz ALV do eksportu
      cl_salv_table=>factory(
        IMPORTING
          r_salv_table = lo_excel
        CHANGING
          t_table      = gt_wydajnosc ).

      " Eksportuj do XLSX
      DATA(lo_result) = lo_excel->to_xml( if_salv_bs_xml=>c_type_xlsx ).
      lv_xstring = lo_result.

      " Konwertuj na solix
      ct_data = cl_bcs_convert=>xstring_to_solix( iv_xstring = lv_xstring ).
      cv_size = xstrlen( lv_xstring ).

    CATCH cx_salv_msg INTO DATA(lx_msg).
      MESSAGE lx_msg TYPE 'E'.
  ENDTRY.
ENDFORM.

*----------------------------------------------------------------------*
* LOCAL CLASS DEFINITION
*----------------------------------------------------------------------*
CLASS lcl_event_handler DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS:
      on_user_command FOR EVENT added_function OF cl_salv_events_table
        IMPORTING e_salv_function,
      on_double_click FOR EVENT double_click OF cl_salv_events_table
        IMPORTING row column.
ENDCLASS.

*----------------------------------------------------------------------*
* LOCAL CLASS IMPLEMENTATION
*----------------------------------------------------------------------*
CLASS lcl_event_handler IMPLEMENTATION.
  METHOD on_user_command.
    CASE e_salv_function.
      WHEN 'HTM_PREVIEW'.
        PERFORM show_htm_preview.
      WHEN 'SEND_HTM'.
        PERFORM send_email USING 'HTM'.
      WHEN 'SEND_XLSX'.
        PERFORM send_email USING 'XLSX'.
    ENDCASE.
  ENDMETHOD.

  METHOD on_double_click.
    DATA: ls_data TYPE ty_wydajnosc.

    READ TABLE gt_wydajnosc INTO ls_data INDEX row.
    IF sy-subrc = 0.
      " Pokaż szczegóły wiersza
      MESSAGE s007(00) WITH 'Wybrano wiersz:' ls_data-pracownik ls_data-produkt.
    ENDIF.
  ENDMETHOD.
ENDCLASS.
