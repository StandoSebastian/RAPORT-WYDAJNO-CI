*&---------------------------------------------------------------------*
*& Report ZREP_WYDAJNOSC_SAP
*&---------------------------------------------------------------------*
*& Raport wydajności SAP dla wydziału produkcyjnego
*& Funkcjonalności:
*&   - Załadowanie danych z plików TSV
*&   - Wyświetlenie danych w ALV
*&   - Podgląd HTML
*&   - Wysyłka maila z załącznikiem HTML lub XLSX
*&---------------------------------------------------------------------*
REPORT zrep_wydajnosc_sap.

*----------------------------------------------------------------------*
* TYPES - Definicje typów danych
*----------------------------------------------------------------------*
TYPES: BEGIN OF ty_wydajnosc,
         wydzial       TYPE char20,      " Wydział produkcyjny
         pracownik     TYPE char50,      " Imię i nazwisko pracownika
         nr_pracownika TYPE char10,      " Numer pracownika
         data          TYPE datum,       " Data
         czas_pracy    TYPE dec_value,   " Czas pracy w godzinach
         ilosc_sztuk   TYPE i,           " Ilość wyprodukowanych sztuk
         norma         TYPE i,           " Norma
         wydajnosc_pct TYPE p DECIMALS 2," Wydajność procentowa
         status        TYPE char10,      " Status (OK/PONIŻEJ/POWYŻEJ)
         uwagi         TYPE char100,     " Uwagi
       END OF ty_wydajnosc.

TYPES: tt_wydajnosc TYPE STANDARD TABLE OF ty_wydajnosc WITH DEFAULT KEY.

*----------------------------------------------------------------------*
* DATA - Deklaracje zmiennych
*----------------------------------------------------------------------*
DATA: gt_wydajnosc TYPE tt_wydajnosc,
      gs_wydajnosc TYPE ty_wydajnosc,
      go_alv       TYPE REF TO cl_salv_table,
      gv_filename  TYPE string,
      gt_file_data TYPE string_table.

*----------------------------------------------------------------------*
* SELECTION SCREEN - Ekran selekcji
*----------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-001.
  PARAMETERS: p_file TYPE rlgrap-filename OBLIGATORY.
  PARAMETERS: p_wydz TYPE char20 DEFAULT 'PRODUKCJA'.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-002.
  PARAMETERS: p_html  RADIOBUTTON GROUP rb1 DEFAULT 'X'.
  PARAMETERS: p_xlsx  RADIOBUTTON GROUP rb1.
SELECTION-SCREEN END OF BLOCK b2.

*----------------------------------------------------------------------*
* AT SELECTION-SCREEN ON VALUE-REQUEST
*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  PERFORM f4_help_file.

*----------------------------------------------------------------------*
* START-OF-SELECTION
*----------------------------------------------------------------------*
START-OF-SELECTION.

  " Załaduj dane z pliku TSV
  PERFORM load_tsv_file.

  " Oblicz wydajność
  PERFORM calculate_wydajnosc.

  " Wyświetl ALV
  PERFORM display_alv.

*&---------------------------------------------------------------------*
*& Form f4_help_file
*&---------------------------------------------------------------------*
*& Pomoc wyszukiwania dla pliku
*&---------------------------------------------------------------------*
FORM f4_help_file.
  DATA: lt_filetable TYPE filetable,
        ls_file      TYPE file_table,
        lv_rc        TYPE i.

  CALL METHOD cl_gui_frontend_services=>file_open_dialog
    EXPORTING
      window_title            = 'Wybierz plik TSV'
      file_filter             = '*.tsv;*.txt'
    CHANGING
      file_table              = lt_filetable
      rc                      = lv_rc
    EXCEPTIONS
      file_open_dialog_failed = 1
      cntl_error              = 2
      error_no_gui            = 3
      not_supported_by_gui    = 4
      OTHERS                  = 5.

  IF sy-subrc = 0 AND lines( lt_filetable ) > 0.
    READ TABLE lt_filetable INTO ls_file INDEX 1.
    p_file = ls_file-filename.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form load_tsv_file
*&---------------------------------------------------------------------*
*& Załadowanie danych z pliku TSV
*&---------------------------------------------------------------------*
FORM load_tsv_file.
  DATA: lt_lines    TYPE string_table,
        lv_line     TYPE string,
        lt_fields   TYPE string_table,
        lv_filename TYPE string.

  lv_filename = p_file.

  " Wczytaj plik z frontendu
  CALL METHOD cl_gui_frontend_services=>gui_upload
    EXPORTING
      filename                = lv_filename
      filetype                = 'ASC'
    CHANGING
      data_tab                = lt_lines
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
      OTHERS                  = 19.

  IF sy-subrc <> 0.
    MESSAGE e001(00) WITH 'Błąd wczytywania pliku TSV'.
    RETURN.
  ENDIF.

  " Przetwórz linie (pomiń nagłówek)
  DATA: lv_first TYPE abap_bool VALUE abap_true.
  LOOP AT lt_lines INTO lv_line.
    IF lv_first = abap_true.
      lv_first = abap_false.
      CONTINUE. " Pomiń nagłówek
    ENDIF.

    CLEAR: gs_wydajnosc, lt_fields.
    SPLIT lv_line AT cl_abap_char_utilities=>horizontal_tab INTO TABLE lt_fields.

    IF lines( lt_fields ) >= 7.
      gs_wydajnosc-wydzial       = VALUE #( lt_fields[ 1 ] OPTIONAL ).
      gs_wydajnosc-pracownik     = VALUE #( lt_fields[ 2 ] OPTIONAL ).
      gs_wydajnosc-nr_pracownika = VALUE #( lt_fields[ 3 ] OPTIONAL ).

      " Konwersja daty
      DATA(lv_date_str) = VALUE #( lt_fields[ 4 ] OPTIONAL ).
      CALL FUNCTION 'CONVERT_DATE_TO_INTERNAL'
        EXPORTING
          date_external = lv_date_str
        IMPORTING
          date_internal = gs_wydajnosc-data
        EXCEPTIONS
          OTHERS        = 1.

      gs_wydajnosc-czas_pracy  = VALUE #( lt_fields[ 5 ] OPTIONAL ).
      gs_wydajnosc-ilosc_sztuk = VALUE #( lt_fields[ 6 ] OPTIONAL ).
      gs_wydajnosc-norma       = VALUE #( lt_fields[ 7 ] OPTIONAL ).

      IF lines( lt_fields ) >= 8.
        gs_wydajnosc-uwagi = VALUE #( lt_fields[ 8 ] OPTIONAL ).
      ENDIF.

      APPEND gs_wydajnosc TO gt_wydajnosc.
    ENDIF.
  ENDLOOP.

  IF gt_wydajnosc IS INITIAL.
    MESSAGE s002(00) WITH 'Brak danych do wyświetlenia' DISPLAY LIKE 'W'.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form calculate_wydajnosc
*&---------------------------------------------------------------------*
*& Obliczenie wydajności procentowej
*&---------------------------------------------------------------------*
FORM calculate_wydajnosc.
  LOOP AT gt_wydajnosc ASSIGNING FIELD-SYMBOL(<fs_wyd>).
    IF <fs_wyd>-norma > 0.
      <fs_wyd>-wydajnosc_pct = ( <fs_wyd>-ilosc_sztuk / <fs_wyd>-norma ) * 100.

      IF <fs_wyd>-wydajnosc_pct >= 100.
        <fs_wyd>-status = 'POWYŻEJ'.
      ELSEIF <fs_wyd>-wydajnosc_pct >= 80.
        <fs_wyd>-status = 'OK'.
      ELSE.
        <fs_wyd>-status = 'PONIŻEJ'.
      ENDIF.
    ELSE.
      <fs_wyd>-wydajnosc_pct = 0.
      <fs_wyd>-status = 'BRAK'.
    ENDIF.
  ENDLOOP.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form display_alv
*&---------------------------------------------------------------------*
*& Wyświetlenie danych w ALV
*&---------------------------------------------------------------------*
FORM display_alv.
  DATA: lo_functions TYPE REF TO cl_salv_functions_list,
        lo_columns   TYPE REF TO cl_salv_columns_table,
        lo_column    TYPE REF TO cl_salv_column_table,
        lo_events    TYPE REF TO cl_salv_events_table.

  TRY.
      cl_salv_table=>factory(
        IMPORTING
          r_salv_table = go_alv
        CHANGING
          t_table      = gt_wydajnosc ).

      " Włącz wszystkie standardowe funkcje
      lo_functions = go_alv->get_functions( ).
      lo_functions->set_all( abap_true ).

      " Konfiguracja kolumn
      lo_columns = go_alv->get_columns( ).
      lo_columns->set_optimize( abap_true ).

      " Ustaw opisy kolumn
      TRY.
          lo_column ?= lo_columns->get_column( 'WYDZIAL' ).
          lo_column->set_short_text( 'Wydział' ).
          lo_column->set_medium_text( 'Wydział prod.' ).
          lo_column->set_long_text( 'Wydział produkcyjny' ).

          lo_column ?= lo_columns->get_column( 'PRACOWNIK' ).
          lo_column->set_short_text( 'Pracownik' ).
          lo_column->set_medium_text( 'Pracownik' ).
          lo_column->set_long_text( 'Imię i nazwisko' ).

          lo_column ?= lo_columns->get_column( 'NR_PRACOWNIKA' ).
          lo_column->set_short_text( 'Nr prac.' ).
          lo_column->set_medium_text( 'Nr pracownika' ).
          lo_column->set_long_text( 'Numer pracownika' ).

          lo_column ?= lo_columns->get_column( 'DATA' ).
          lo_column->set_short_text( 'Data' ).
          lo_column->set_medium_text( 'Data' ).
          lo_column->set_long_text( 'Data pracy' ).

          lo_column ?= lo_columns->get_column( 'CZAS_PRACY' ).
          lo_column->set_short_text( 'Czas' ).
          lo_column->set_medium_text( 'Czas pracy' ).
          lo_column->set_long_text( 'Czas pracy [h]' ).

          lo_column ?= lo_columns->get_column( 'ILOSC_SZTUK' ).
          lo_column->set_short_text( 'Ilość' ).
          lo_column->set_medium_text( 'Ilość sztuk' ).
          lo_column->set_long_text( 'Ilość wyprodukowanych sztuk' ).

          lo_column ?= lo_columns->get_column( 'NORMA' ).
          lo_column->set_short_text( 'Norma' ).
          lo_column->set_medium_text( 'Norma' ).
          lo_column->set_long_text( 'Norma produkcyjna' ).

          lo_column ?= lo_columns->get_column( 'WYDAJNOSC_PCT' ).
          lo_column->set_short_text( 'Wydaj.%' ).
          lo_column->set_medium_text( 'Wydajność %' ).
          lo_column->set_long_text( 'Wydajność procentowa' ).

          lo_column ?= lo_columns->get_column( 'STATUS' ).
          lo_column->set_short_text( 'Status' ).
          lo_column->set_medium_text( 'Status' ).
          lo_column->set_long_text( 'Status wydajności' ).

          lo_column ?= lo_columns->get_column( 'UWAGI' ).
          lo_column->set_short_text( 'Uwagi' ).
          lo_column->set_medium_text( 'Uwagi' ).
          lo_column->set_long_text( 'Uwagi dodatkowe' ).

        CATCH cx_salv_not_found.
          " Ignoruj błędy kolumn
      ENDTRY.

      " Ustaw tytuł
      DATA(lo_display) = go_alv->get_display_settings( ).
      lo_display->set_striped_pattern( abap_true ).
      lo_display->set_list_header( |Raport wydajności - { p_wydz }| ).

      " Zarejestruj eventy dla przycisków
      SET HANDLER lcl_event_handler=>on_user_command FOR go_alv->get_event( ).

      " Dodaj własne przyciski
      PERFORM add_custom_buttons.

      " Wyświetl ALV
      go_alv->display( ).

    CATCH cx_salv_msg INTO DATA(lx_msg).
      MESSAGE lx_msg TYPE 'E'.
  ENDTRY.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form add_custom_buttons
*&---------------------------------------------------------------------*
*& Dodanie własnych przycisków do toolbara ALV
*&---------------------------------------------------------------------*
FORM add_custom_buttons.
  DATA: lo_functions TYPE REF TO cl_salv_functions_list.

  lo_functions = go_alv->get_functions( ).

  TRY.
      lo_functions->add_function(
        name     = 'ZHTML'
        icon     = '@0S@'
        text     = 'Podgląd HTML'
        tooltip  = 'Wyświetl podgląd HTML'
        position = if_salv_c_function_position=>right_of_salv_functions ).

      lo_functions->add_function(
        name     = 'ZMAIL'
        icon     = '@3Y@'
        text     = 'Wyślij email'
        tooltip  = 'Wyślij raport emailem'
        position = if_salv_c_function_position=>right_of_salv_functions ).

    CATCH cx_salv_existing cx_salv_wrong_call.
      " Ignoruj błędy
  ENDTRY.
ENDFORM.

*----------------------------------------------------------------------*
* CLASS lcl_event_handler DEFINITION
*----------------------------------------------------------------------*
CLASS lcl_event_handler DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS: on_user_command FOR EVENT added_function OF cl_salv_events
      IMPORTING e_salv_function.
ENDCLASS.

*----------------------------------------------------------------------*
* CLASS lcl_event_handler IMPLEMENTATION
*----------------------------------------------------------------------*
CLASS lcl_event_handler IMPLEMENTATION.
  METHOD on_user_command.
    CASE e_salv_function.
      WHEN 'ZHTML'.
        PERFORM show_html_preview.
      WHEN 'ZMAIL'.
        PERFORM send_email.
    ENDCASE.
  ENDMETHOD.
ENDCLASS.

*&---------------------------------------------------------------------*
*& Form show_html_preview
*&---------------------------------------------------------------------*
*& Wyświetlenie podglądu HTML
*&---------------------------------------------------------------------*
FORM show_html_preview.
  DATA: lv_html    TYPE string,
        lo_browser TYPE REF TO cl_gui_html_viewer,
        lt_html    TYPE TABLE OF w3html.

  " Generuj HTML
  PERFORM generate_html CHANGING lv_html.

  " Konwertuj do tabeli dla przeglądarki
  DATA: lv_len  TYPE i,
        lv_off  TYPE i VALUE 0,
        lv_line TYPE w3html.

  lv_len = strlen( lv_html ).
  WHILE lv_off < lv_len.
    IF lv_len - lv_off >= 255.
      lv_line = lv_html+lv_off(255).
    ELSE.
      lv_line = lv_html+lv_off.
    ENDIF.
    APPEND lv_line TO lt_html.
    lv_off = lv_off + 255.
  ENDWHILE.

  " Wyświetl w nowym oknie
  DATA: lv_url TYPE w3url.

  CALL FUNCTION 'DP_CREATE_URL'
    EXPORTING
      type     = 'text'
      subtype  = 'html'
    TABLES
      data     = lt_html
    CHANGING
      url      = lv_url
    EXCEPTIONS
      OTHERS   = 1.

  IF sy-subrc = 0.
    CALL FUNCTION 'CALL_BROWSER'
      EXPORTING
        url                    = lv_url
      EXCEPTIONS
        frontend_not_supported = 1
        frontend_error         = 2
        prog_not_found         = 3
        no_batch               = 4
        unspecified_error      = 5
        OTHERS                 = 6.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form generate_html
*&---------------------------------------------------------------------*
*& Generowanie kodu HTML z danymi
*&---------------------------------------------------------------------*
FORM generate_html CHANGING cv_html TYPE string.
  DATA: lv_rows TYPE string,
        lv_date TYPE string.

  " Nagłówek HTML
  cv_html = |<!DOCTYPE html>| &&
            |<html lang="pl">| &&
            |<head>| &&
            |<meta charset="UTF-8">| &&
            |<title>Raport Wydajności - { p_wydz }</title>| &&
            |<style>| &&
            |body \{ font-family: Arial, sans-serif; margin: 20px; \}| &&
            |h1 \{ color: #333; \}| &&
            |table \{ border-collapse: collapse; width: 100%; \}| &&
            |th, td \{ border: 1px solid #ddd; padding: 8px; text-align: left; \}| &&
            |th \{ background-color: #4CAF50; color: white; \}| &&
            |tr:nth-child(even) \{ background-color: #f2f2f2; \}| &&
            |.status-ok \{ color: green; \}| &&
            |.status-powyzej \{ color: blue; font-weight: bold; \}| &&
            |.status-ponizej \{ color: red; \}| &&
            |</style>| &&
            |</head>| &&
            |<body>| &&
            |<h1>Raport Wydajności - { p_wydz }</h1>| &&
            |<p>Data wygenerowania: { sy-datum DATE = USER }</p>| &&
            |<table>| &&
            |<tr>| &&
            |<th>Wydział</th>| &&
            |<th>Pracownik</th>| &&
            |<th>Nr</th>| &&
            |<th>Data</th>| &&
            |<th>Czas [h]</th>| &&
            |<th>Ilość</th>| &&
            |<th>Norma</th>| &&
            |<th>Wydajność %</th>| &&
            |<th>Status</th>| &&
            |<th>Uwagi</th>| &&
            |</tr>|.

  " Wiersze danych
  LOOP AT gt_wydajnosc INTO gs_wydajnosc.
    WRITE gs_wydajnosc-data TO lv_date DD/MM/YYYY.

    DATA(lv_status_class) = SWITCH string( gs_wydajnosc-status
      WHEN 'OK'      THEN 'status-ok'
      WHEN 'POWYŻEJ' THEN 'status-powyzej'
      WHEN 'PONIŻEJ' THEN 'status-ponizej'
      ELSE '' ).

    cv_html = cv_html &&
      |<tr>| &&
      |<td>{ gs_wydajnosc-wydzial }</td>| &&
      |<td>{ gs_wydajnosc-pracownik }</td>| &&
      |<td>{ gs_wydajnosc-nr_pracownika }</td>| &&
      |<td>{ lv_date }</td>| &&
      |<td>{ gs_wydajnosc-czas_pracy }</td>| &&
      |<td>{ gs_wydajnosc-ilosc_sztuk }</td>| &&
      |<td>{ gs_wydajnosc-norma }</td>| &&
      |<td>{ gs_wydajnosc-wydajnosc_pct }</td>| &&
      |<td class="{ lv_status_class }">{ gs_wydajnosc-status }</td>| &&
      |<td>{ gs_wydajnosc-uwagi }</td>| &&
      |</tr>|.
  ENDLOOP.

  " Zamknięcie HTML
  cv_html = cv_html &&
            |</table>| &&
            |</body>| &&
            |</html>|.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form send_email
*&---------------------------------------------------------------------*
*& Wysyłanie emaila z raportem
*&---------------------------------------------------------------------*
FORM send_email.
  DATA: lv_email   TYPE ad_smtpadr,
        lv_subject TYPE so_obj_des,
        lv_html    TYPE string.

  " Pobierz adres email
  CALL FUNCTION 'POPUP_TO_GET_VALUE'
    EXPORTING
      fieldname  = 'AD_SMTPADR'
      tabname    = 'ADR6'
      titel      = 'Podaj adres email'
      valuein    = ''
    IMPORTING
      valueout   = lv_email
    EXCEPTIONS
      OTHERS     = 1.

  IF sy-subrc <> 0 OR lv_email IS INITIAL.
    MESSAGE s003(00) WITH 'Anulowano wysyłkę emaila' DISPLAY LIKE 'W'.
    RETURN.
  ENDIF.

  lv_subject = |Raport wydajności - { p_wydz } - { sy-datum }|.

  " Sprawdź typ załącznika
  IF p_html = abap_true.
    PERFORM send_email_html USING lv_email lv_subject.
  ELSE.
    PERFORM send_email_xlsx USING lv_email lv_subject.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form send_email_html
*&---------------------------------------------------------------------*
*& Wysyłanie emaila z załącznikiem HTML
*&---------------------------------------------------------------------*
FORM send_email_html USING iv_email   TYPE ad_smtpadr
                           iv_subject TYPE so_obj_des.
  DATA: lo_mail       TYPE REF TO cl_bcs,
        lo_doc        TYPE REF TO cl_document_bcs,
        lo_sender     TYPE REF TO cl_sapuser_bcs,
        lo_recipient  TYPE REF TO if_recipient_bcs,
        lt_html       TYPE soli_tab,
        lv_html       TYPE string,
        lv_filename   TYPE sood-objdes.

  TRY.
      " Utwórz obiekt maila
      lo_mail = cl_bcs=>create_persistent( ).

      " Generuj HTML
      PERFORM generate_html CHANGING lv_html.

      " Konwertuj HTML do tabeli
      DATA(lo_conv) = cl_bcs_convert=>string_to_soli( lv_html ).
      lt_html = lo_conv.

      " Utwórz dokument
      lo_doc = cl_document_bcs=>create_document(
        i_type    = 'HTM'
        i_text    = lt_html
        i_subject = iv_subject ).

      " Dodaj załącznik HTML
      lv_filename = |Raport_wydajnosci_{ sy-datum }.html|.
      lo_doc->add_attachment(
        i_attachment_type    = 'HTM'
        i_attachment_subject = lv_filename
        i_att_content_text   = lt_html ).

      " Ustaw dokument
      lo_mail->set_document( lo_doc ).

      " Ustaw nadawcę
      lo_sender = cl_sapuser_bcs=>create( sy-uname ).
      lo_mail->set_sender( lo_sender ).

      " Dodaj odbiorcę
      lo_recipient = cl_cam_address_bcs=>create_internet_address( iv_email ).
      lo_mail->add_recipient(
        i_recipient  = lo_recipient
        i_express    = abap_true ).

      " Wyślij
      lo_mail->send( i_with_error_screen = abap_true ).
      COMMIT WORK.

      MESSAGE s004(00) WITH 'Email został wysłany pomyślnie'.

    CATCH cx_bcs INTO DATA(lx_bcs).
      MESSAGE lx_bcs TYPE 'E'.
  ENDTRY.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form send_email_xlsx
*&---------------------------------------------------------------------*
*& Wysyłanie emaila z załącznikiem XLSX
*&---------------------------------------------------------------------*
FORM send_email_xlsx USING iv_email   TYPE ad_smtpadr
                           iv_subject TYPE so_obj_des.
  DATA: lo_mail       TYPE REF TO cl_bcs,
        lo_doc        TYPE REF TO cl_document_bcs,
        lo_sender     TYPE REF TO cl_sapuser_bcs,
        lo_recipient  TYPE REF TO if_recipient_bcs,
        lt_body       TYPE soli_tab,
        ls_body       TYPE soli,
        lt_xlsx       TYPE solix_tab,
        lv_size       TYPE i,
        lv_filename   TYPE sood-objdes.

  TRY.
      " Utwórz obiekt maila
      lo_mail = cl_bcs=>create_persistent( ).

      " Przygotuj treść wiadomości
      ls_body-line = |Szanowni Państwo,|.
      APPEND ls_body TO lt_body.
      CLEAR ls_body.
      APPEND ls_body TO lt_body.
      ls_body-line = |W załączeniu raport wydajności dla wydziału { p_wydz }.|.
      APPEND ls_body TO lt_body.
      CLEAR ls_body.
      APPEND ls_body TO lt_body.
      ls_body-line = |Z poważaniem,|.
      APPEND ls_body TO lt_body.
      ls_body-line = |System SAP|.
      APPEND ls_body TO lt_body.

      " Konwertuj dane do XLSX
      PERFORM convert_to_xlsx CHANGING lt_xlsx lv_size.

      " Utwórz dokument
      lo_doc = cl_document_bcs=>create_document(
        i_type    = 'RAW'
        i_text    = lt_body
        i_subject = iv_subject ).

      " Dodaj załącznik XLSX
      lv_filename = |Raport_wydajnosci_{ sy-datum }.xlsx|.
      lo_doc->add_attachment(
        i_attachment_type    = 'XLS'
        i_attachment_subject = lv_filename
        i_attachment_size    = lv_size
        i_att_content_hex    = lt_xlsx ).

      " Ustaw dokument
      lo_mail->set_document( lo_doc ).

      " Ustaw nadawcę
      lo_sender = cl_sapuser_bcs=>create( sy-uname ).
      lo_mail->set_sender( lo_sender ).

      " Dodaj odbiorcę
      lo_recipient = cl_cam_address_bcs=>create_internet_address( iv_email ).
      lo_mail->add_recipient(
        i_recipient  = lo_recipient
        i_express    = abap_true ).

      " Wyślij
      lo_mail->send( i_with_error_screen = abap_true ).
      COMMIT WORK.

      MESSAGE s004(00) WITH 'Email został wysłany pomyślnie'.

    CATCH cx_bcs INTO DATA(lx_bcs).
      MESSAGE lx_bcs TYPE 'E'.
  ENDTRY.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form convert_to_xlsx
*&---------------------------------------------------------------------*
*& Konwersja danych do formatu XLSX
*&---------------------------------------------------------------------*
FORM convert_to_xlsx CHANGING ct_xlsx TYPE solix_tab
                              cv_size TYPE i.
  DATA: lo_result TYPE REF TO cl_salv_ex_result_data_table,
        lt_table  TYPE REF TO data.

  TRY.
      " Użyj standardowej konwersji ALV do XLSX
      CREATE DATA lt_table LIKE gt_wydajnosc.
      ASSIGN lt_table->* TO FIELD-SYMBOL(<lt_data>).
      <lt_data> = gt_wydajnosc.

      cl_salv_table=>factory(
        IMPORTING
          r_salv_table = DATA(lo_salv)
        CHANGING
          t_table      = <lt_data> ).

      DATA(lo_res) = cl_salv_bs_lex=>export_from_salv_table(
        EXPORTING
          is_call_params = VALUE cl_salv_bs_lex=>s_call_params( xml_flavour = if_salv_bs_xml=>c_flavour_xlsx )
          r_salv_table   = lo_salv ).

      " Konwertuj xstring do solix
      DATA(lv_xstring) = lo_res->get_data( ).
      cv_size = xstrlen( lv_xstring ).

      ct_xlsx = cl_bcs_convert=>xstring_to_solix( lv_xstring ).

    CATCH cx_root INTO DATA(lx_root).
      MESSAGE lx_root TYPE 'E'.
  ENDTRY.
ENDFORM.
