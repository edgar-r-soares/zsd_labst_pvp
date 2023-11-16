*&---------------------------------------------------------------------*
*& Report  YEDGAR_LABST_PVP
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report  zsd_labst_pvp.

tables: kna1, knb1, mean, A154, mara, t001w, mbew, marc, mcha.

types: begin of t_data,
         vkorg type A154-vkorg,
         vtweg type A154-vtweg,
         pltyp type A154-pltyp,
         katr6 type kna1-katr6,
         altkn type knb1-altkn,
         werks type t001w-werks,
         name1 type t001w-name1,
         ean11 type mean-ean11,
         matnr type mara-matnr,
         maktx type makt-maktx,
         lbkum type mbew-lbkum,
         meinh type mean-meinh,
         kbetr type vbap-netpr, "konp-kbetr
         konwa type konp-konwa,
         vfdat type mcha-vfdat,
         xchpf type marc-xchpf,
         meins type mara-meins,
         lvorm type marc-lvorm,
         taxm1 type a002-taxm1,
         kbetr_tax type konp-kbetr,
         konwa_tax type konp-konwa,
       end of t_data.

data: gt_data type table of t_data,
      gr_salv type ref to cl_salv_table.

select-options: s_werks for t001w-werks,
                s_lvorm for marc-lvorm default space sign i option eq,
                s_katr6 for kna1-katr6,
                s_altkn for knb1-altkn,
                s_matkl for mara-matkl,
                s_matnr for mara-matnr,
                s_ean11 for mean-ean11,
                s_eantp for mean-eantp default 'ZK',
                s_lbkum for mbew-lbkum default 0 sign i option gt,
                s_vkorg for A154-vkorg default '1110',
                s_vtweg for A154-vtweg,
                s_pltyp for A154-pltyp default 'K1',
                s_vfdat for mcha-VFDAT.
parameters: p_date type vbkd-prsdt obligatory default sy-datum.


selection-screen begin of block b6 with frame title text-006.
parameters p_var type slis_vari.
selection-screen end of block b6.

selection-screen begin of block b1 with frame title text-001.
parameters: p_sender type adr6-smtp_addr default 'infosys@tba.co.ao' obligatory.
parameters: p_title type so_obj_des.
selection-screen begin of block b2 with frame title text-002.
parameters: p_email  type adr6-smtp_addr." DEFAULT 'edgar.soares@everedge.pt',
selection-screen end of block b2.
selection-screen begin of block b3 with frame title text-003.
parameters: p_attach as checkbox.
selection-screen end of block b3.
selection-screen end of block b1.


start-of-selection.
  perform f_select_data.


end-of-selection.
  perform f_display_data.



*&---------------------------------------------------------------------*
*&      Form  f_send_email
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
form f_send_email .
  data: lt_head type bcsy_text.
  data: lt_foot type bcsy_text,
        mail_title type so_obj_des.
  data: l_date(10).
  field-symbols <ls_html> type bcsy_text.

  mail_title = p_title.
  replace 'SYSID' in mail_title with sy-sysid.
  write p_date to l_date.
  replace 'DATUM' in mail_title with l_date.

  append 'Bom dia,' to lt_head.
  append initial line to lt_head.
  append 'Segue a listagem de PVPs:' to lt_head.

  append initial line to lt_foot.
  append 'Obrigado,' to lt_foot.
  append 'SAP' to lt_foot.

  call method zsend_email=>salv_table
    exporting
      ir_salv       = gr_salv
      it_table      = gt_data
      i_sender      = p_sender
      i_receiver    = p_email
      i_mail_title  = mail_title
      i_attach      = p_attach
      i_filename    = 'PVPs e Stocks.XLSX'
      it_header     = lt_head
      it_footer     = lt_foot
      i_sheet_title = 'PVP'.
endform.                    "f_send_email
*&---------------------------------------------------------------------*
*&      Form  F_SELECT_DATA
*&---------------------------------------------------------------------*
form f_select_data .
  field-symbols: <ls_data> type t_data.

  select A154~vkorg
         A154~vtweg
         A154~pltyp
         kna1~katr6
         knb1~altkn
         t001w~werks
         t001w~name1
         mean~ean11
         mara~matnr
         makt~maktx
         mbew~lbkum
         mean~meinh
         konp~kbetr
         konp~konwa
*         zmarc_expirydate~vfdat
         marc~xchpf
         mara~meins
         marc~lvorm
         a002~taxm1
         konp2~kbetr as kbetr_tax
         konp2~konwa as konwa_tax
         into CORRESPONDING FIELDS OF table gt_data
         from mbew
         join mara on mara~matnr eq mbew~matnr
         join makt on makt~matnr eq mbew~matnr
         join mean on mean~matnr eq mbew~matnr "AND mean~meinh EQ mara~meins
         join t001k on t001k~bwkey eq mbew~bwkey
         join t001w on t001w~bwkey eq t001k~bwkey
         join marc on marc~matnr eq mbew~matnr and marc~werks eq t001w~werks
         join kna1 on kna1~kunnr eq t001w~kunnr
         join knb1 on knb1~kunnr eq t001w~kunnr and knb1~bukrs eq t001k~bukrs
         join A154 on A154~matnr eq mara~matnr "AND A154~vkorg EQ t001w~vkorg AND A154~vtweg EQ t001w~vtweg
          and A154~vrkme eq mean~meinh
         join konp on konp~knumh eq A154~knumh
         join tvko on tvko~bukrs eq t001k~bukrs and tvko~vkorg eq A154~vkorg
         join tvkwz on tvkwz~werks eq t001w~werks and tvkwz~vkorg eq A154~vkorg and tvkwz~vtweg eq A154~vtweg
         join mlan on mlan~matnr eq mara~matnr and mlan~aland eq t001w~land1
         join knvi on knvi~kunnr eq kna1~kunnr and knvi~aland eq kna1~land1
         join a002 on a002~taxk1 eq knvi~taxkd and a002~taxm1 eq mlan~taxm1 and a002~aland eq t001w~land1
         join konp as konp2 on konp2~knumh eq a002~knumh
*         left join zmarc_expirydate on zmarc_expirydate~matnr eq marc~matnr and zmarc_expirydate~werks eq marc~werks
         where A154~kappl eq 'V' and A154~kschl eq 'VKP0' and A154~vkorg in s_vkorg and A154~vtweg in s_vtweg and A154~pltyp in s_pltyp
           and mbew~matnr in s_matnr and mbew~lbkum in s_lbkum
           and kna1~katr6 in s_katr6
           and knb1~altkn in s_altkn
           and t001w~werks in s_werks
           and marc~lvorm in s_lvorm
           and mara~matkl in s_matkl
           and mean~ean11 in s_ean11
           and mean~hpean eq 'X'
           and mean~eantp in s_eantp
*           and A154~kfrst eq space
           and A154~datbi ge p_date and A154~datab le p_date
           and knvi~tatyp eq 'ZWST'
           and a002~kappl eq 'V' and a002~kschl eq 'MWST' and a002~datbi ge p_date and a002~datab le p_date.
  delete gt_data where vfdat not in s_vfdat.

  loop at gt_data assigning <ls_data>.
    divide <ls_data>-kbetr_tax by 10.
    if <ls_data>-konwa_tax is initial.
      <ls_data>-konwa_tax = '%'.
    endif.
    select single maktm into <ls_data>-maktx from mamt where matnr eq <ls_data>-matnr and meinh eq <ls_data>-meinh and spras eq sy-langu.
    if <ls_data>-meins ne <ls_data>-meinh.
      clear: <ls_data>-lbkum.
    endif.
    if <ls_data>-xchpf is not initial.
      select min( vfdat ) into <ls_data>-vfdat from mcha join mchb on mchb~charg eq mcha~charg and mchb~matnr eq mchb~matnr and mchb~werks eq mchb~werks
             where mchb~matnr eq <ls_data>-matnr and mchb~werks eq <ls_data>-werks and clabs gt 0.
    endif.
  endloop.

endform.                    " F_SELECT_DATA
*&---------------------------------------------------------------------*
*&      Form  F_DISPLAY_DATA
*&---------------------------------------------------------------------*
form f_display_data .
  cl_salv_table=>factory(
    importing
      r_salv_table = gr_salv
    changing
      t_table      = gt_data ).

  data: lr_functions type ref to cl_salv_functions_list.
  lr_functions = gr_salv->get_functions( ).
  lr_functions->set_all( abap_true ).
  lr_functions->set_default( abap_true ).
  data: lo_funcs type salv_t_ui_func.
  data: lo_func type salv_s_ui_func.
  lo_funcs = lr_functions->get_functions( ).
  loop at lo_funcs into lo_func.
    lo_func-r_function->set_enable( value = 'X' ).
    lo_func-r_function->set_visible( value = 'X' ).
  endloop.

  data: lo_layout  type ref to cl_salv_layout,
        lf_variant type slis_vari,
        ls_key    type salv_s_layout_key.
  lo_layout = gr_salv->get_layout( ).
  ls_key-report = sy-repid.
  lo_layout->set_key( ls_key ).
  lo_layout->set_save_restriction( if_salv_c_layout=>restrict_none ).
  lo_layout->set_default( abap_true ).

  lf_variant = p_var.

  lo_layout->set_initial_layout( lf_variant ).


  data: lo_column      type ref to cl_salv_column.
  data: lo_columns     type ref to cl_salv_columns_table.
  data: lo_col     type ref to cl_salv_column_table.

  try.
      lo_columns = gr_salv->get_columns( ).

      lo_col ?= lo_columns->get_column( 'LBKUM' ).
      lo_col->set_quantity_column( 'MEINH' ).
      lo_col ?= lo_columns->get_column( 'KBETR' ).
      lo_col->set_currency_column( 'KONWA' ).

      lo_col ?= lo_columns->get_column( 'KBETR_TAX' ).
      lo_col->set_currency_column( 'KONWA_TAX' ).

      lo_col ?= lo_columns->get_column( 'LVORM' ).
      lo_col->set_cell_type( if_salv_c_cell_type=>checkbox ).

      lo_column = lo_columns->get_column( 'KBETR_TAX' ).
      lo_column->set_short_text( 'IVA' ).
      lo_column->set_medium_text( 'IVA' ).
      lo_column->set_long_text( 'IVA' ).

*      lo_col ?= lo_columns->get_column( 'UMLME' ).
*      lo_col->set_quantity_column( 'MEINS' ).
*      lo_col ?= lo_columns->get_column( 'INSME' ).
*      lo_col->set_quantity_column( 'MEINS' ).
*      lo_col ?= lo_columns->get_column( 'EINME' ).
*      lo_col->set_quantity_column( 'MEINS' ).
*      lo_col ?= lo_columns->get_column( 'SPEME' ).
*      lo_col->set_quantity_column( 'MEINS' ).
*      lo_column->set_medium_text( 'Mensagem' ).
*      lo_column->set_long_text( 'Mensagem' ).
*      lo_column = lo_columns->get_column( 'MSGAC' ).

    catch cx_salv_not_found.                            "#EC NO_HANDLER
  endtry.


  if p_email is not initial.
    perform f_send_email.
  else.
    gr_salv->display( ).
  endif.
endform.                    " F_DISPLAY_DATA
