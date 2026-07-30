from templategate.core.selector import match_attributes, match_selector


def test_wildcard_matches_everything():
    assert match_selector("*", "Sheet1!B2")
    assert match_selector("*", "sheet:Sheet2")
    assert match_selector("*", "p12")


def test_excel_range_containment():
    assert match_selector("Sheet1!B2:B100", "Sheet1!B2")
    assert match_selector("Sheet1!B2:B100", "Sheet1!B100")
    assert not match_selector("Sheet1!B2:B100", "Sheet1!C2")
    assert not match_selector("Sheet1!B2:B100", "Sheet2!B2")


def test_japanese_sheet_names():
    assert match_selector("計画表!B4:B6", "計画表!B5")
    assert not match_selector("計画表!B4:B6", "計画表!B7")
    assert match_selector("計画表", "計画表#print")
    assert match_selector("計画表", "計画表!Z99")


def test_whole_sheet_selector_does_not_match_other_namespaces():
    assert not match_selector("Sheet1", "sheet:Sheet1")
    assert not match_selector("Sheet1", "vba")
    assert match_selector("sheet:*", "sheet:Sheet1")
    assert match_selector("vba", "vba")


def test_ranged_selector_does_not_match_sheet_level_locations():
    assert not match_selector("Sheet1!A1:Z100", "Sheet1#print")
    assert not match_selector("Sheet1!A1:Z100", "Sheet1#image:abcd1234")


def test_word_selectors():
    assert match_selector("body", "p3")
    assert not match_selector("body", "table1!r1c1")
    assert match_selector("p3-10", "p7")
    assert not match_selector("p3-10", "p11")
    assert match_selector("table2", "table2!r1c3")
    assert not match_selector("table2", "table3!r1c3")
    assert match_selector("section1", "section1#header_footer")


def test_whole_column_and_row_selectors():
    """Open-ended ranges report None boundaries and used to raise TypeError."""
    assert match_selector("Sheet1!B:B", "Sheet1!B5")
    assert match_selector("Sheet1!B:B", "Sheet1!B1048576")
    assert not match_selector("Sheet1!B:B", "Sheet1!C5")
    assert match_selector("Sheet1!A:Z", "Sheet1!C3")
    assert not match_selector("Sheet1!A:Z", "Sheet1!AA3")
    assert match_selector("Sheet1!4:4", "Sheet1!B4")
    assert not match_selector("Sheet1!4:4", "Sheet1!B5")


def test_sheet_names_containing_the_separator_characters():
    assert match_selector("Q1#Draft!B2:B5", "Q1#Draft!B3")
    assert not match_selector("Q1#Draft!B2:B5", "Q1#Draft!C3")
    assert match_selector("Q1#Draft", "Q1#Draft!B3")
    assert match_selector("Q1#Draft", "Q1#Draft#print")
    assert not match_selector("Q1#Draft!B2:B5", "Q1#Draft#print")
    assert match_selector("Hi!There", "Hi!There!B2")


def test_sheet_level_kind_selectors():
    assert match_selector("Sheet1#print", "Sheet1#print")
    assert not match_selector("Sheet1#print", "Sheet1#header_footer")
    assert not match_selector("Sheet1#print", "Sheet1!B2")


def test_attribute_matching():
    assert match_attributes(["*"], "value")
    assert match_attributes(["value", "formula"], "value")
    assert not match_attributes(["value"], "format")
