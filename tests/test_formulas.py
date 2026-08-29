from ledgerflow.formulas import shift_for_insert, shift_range_ref, translate_rows

SAME_SHEET = dict(sheet="Data", target_sheet="Data", at_row=12, count=3)


def test_range_ending_on_the_last_data_row_grows():
    assert shift_for_insert("=SUM(D4:D11)", **SAME_SHEET) == "=SUM(D4:D14)"


def test_range_spanning_the_insertion_point_grows():
    assert shift_for_insert("=SUM(D4:D100)", **SAME_SHEET) == "=SUM(D4:D103)"


def test_range_entirely_above_is_untouched():
    assert shift_for_insert("=SUM(D4:D9)", **SAME_SHEET) == "=SUM(D4:D9)"


def test_reference_below_the_insertion_moves_down():
    assert shift_for_insert("=B13*2", **SAME_SHEET) == "=B16*2"


def test_absolute_rows_still_grow_at_the_range_end():
    assert shift_for_insert("=SUM($D$4:$D$11)", **SAME_SHEET) == "=SUM($D$4:$D$14)"


def test_other_sheets_are_left_alone():
    assert shift_for_insert("=Summary!B11", **SAME_SHEET) == "=Summary!B11"


def test_cross_sheet_reference_to_the_target_grows():
    assert (
        shift_for_insert("=SUM(Data!D4:D11)", sheet="Summary", target_sheet="Data", at_row=12, count=3)
        == "=SUM(Data!D4:D14)"
    )


def test_quoted_sheet_names_are_handled():
    assert (
        shift_for_insert("=SUM('My Ledger'!D4:D11)", sheet="S", target_sheet="My Ledger", at_row=12, count=3)
        == "=SUM('My Ledger'!D4:D14)"
    )


def test_text_inside_a_string_literal_is_never_rewritten():
    assert shift_for_insert('=IF(B4="D11 ref",D11,0)', **SAME_SHEET) == '=IF(B4="D11 ref",D11,0)'


def test_function_names_containing_a_cell_shape_survive():
    assert shift_for_insert("=LOG10(D13)", **SAME_SHEET) == "=LOG10(D16)"


def test_fill_down_shifts_relative_rows_only():
    assert translate_rows("=D11-E11", 3) == "=D14-E14"
    assert translate_rows("=D11*$B$1", 3) == "=D14*$B$1"


def test_plain_range_strings_shift_too():
    assert shift_range_ref("A3:G11", at_row=12, count=3) == "A3:G14"
