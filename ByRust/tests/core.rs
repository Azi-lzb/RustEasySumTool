use bigdecimal::{BigDecimal, Zero};
use excel_summary::{
    functions::FunctionRegistry,
    model::Severity,
    numbers::{divide, parse_text},
    rules::{parse_reference, parse_rule},
};
use std::collections::BTreeMap;

fn bindings() -> BTreeMap<String, String> {
    [("FUN_1", "SUM"), ("FUN_2", "WAVG"), ("FUN_99", "SUM")]
        .map(|(a, b)| (a.into(), b.into()))
        .into()
}
fn parsed(text: &str, sheet: &str, cell: &str) -> excel_summary::rules::TemplateRule {
    parse_rule(text, sheet, cell, &bindings(), &FunctionRegistry::default())
        .unwrap()
        .unwrap()
}

#[test]
fn omitted_slots_preserve_order_and_duplicates() {
    for (text, expected) in [
        ("FUN_1", vec!["K4"]),
        ("fun_1()", vec!["K4"]),
        ("FUN_1(   )", vec!["K4"]),
        ("FUN_1(,B5)", vec!["K4", "B5"]),
        ("FUN_1(A5,)", vec!["A5", "K4"]),
        ("FUN_1(,)", vec!["K4", "K4"]),
        ("FUN_1(A5,,C5)", vec!["A5", "K4", "C5"]),
        ("FUN_1(,,)", vec!["K4", "K4", "K4"]),
        ("FUN_1(A1:A3,)", vec!["A1:A3", "K4"]),
    ] {
        let rule = parsed(text, "表", "K4");
        assert_eq!(
            rule.references
                .iter()
                .map(|r| r.coordinate.as_str())
                .collect::<Vec<_>>(),
            expected,
            "{text}"
        );
    }
}
#[test]
fn defaults_use_each_template_location_and_sheet() {
    for (sheet, cell) in [("本表", "D8"), ("另表", "XFD1048576")] {
        let rule = parsed("FUN_1('显式,表'!A1,)", sheet, cell);
        assert_eq!(rule.references[1].sheet, sheet);
        assert_eq!(rule.references[1].coordinate, cell);
        assert_eq!(rule.references[0].sheet, "显式,表");
    }
}
#[test]
fn quote_escaping_and_fullwidth_are_supported() {
    for text in [
        "wavg（'O''Brien,表'!$A$1:$A$3，B1:B3）",
        "WAVG('O''Brien,表'!A1:A3,B1:B3）",
    ] {
        let rule = parsed(text, "本表", "D8");
        assert_eq!(rule.references.len(), 2);
        assert_eq!(rule.references[0].sheet, "O'Brien,表");
        assert_eq!(rule.references[0].shape(), (3, 1));
    }
    assert_eq!(
        parsed("SUM（明细（甲）!A1）", "表", "D8").references[0].sheet,
        "明细（甲）"
    );
}
#[test]
fn alias_identifiers_never_imply_arity() {
    for text in ["F2_1(A1,B1)", "FUN_2(A1,B1)"] {
        assert_eq!(parsed(text, "表", "K4").operation, "WAVG");
    }
    assert_eq!(parsed("FUN_99(A1,A2,A3)", "表", "K4").references.len(), 3);
    for text in ["fun_2()", "FUN_2(A1,B1,C1)"] {
        let error = parse_rule(
            text,
            "模板",
            "K4",
            &bindings(),
            &FunctionRegistry::default(),
        )
        .unwrap_err();
        for fragment in ["模板", "K4", text, "FUN_2", "WAVG", "期望 2", "实际"] {
            assert!(error.contains(fragment), "{error}");
        }
    }
    let custom = BTreeMap::from([
        ("F2_1".into(), "SUM".into()),
        ("F1_1".into(), "WAVG".into()),
    ]);
    assert_eq!(
        parse_rule(
            "F2_1(,,)",
            "表",
            "B9",
            &custom,
            &FunctionRegistry::default()
        )
        .unwrap()
        .unwrap()
        .references
        .len(),
        3
    );
}
#[test]
fn shape_is_not_flattened_or_broadcast() {
    for text in [
        "WAVG(A1:A3,B1:D1)",
        "WAVG(A1:A2,B1)",
        "AVG(A1:A2)",
        "SUMPRODUCT(A1:A2,B1:B2)",
        "RATIO(A1:A2,B1:B2)",
    ] {
        assert!(
            parse_rule(text, "表", "D8", &bindings(), &FunctionRegistry::default()).is_err(),
            "{text}"
        );
    }
    let reference = parse_reference("$B$2:$D$3", "表").unwrap();
    assert_eq!(
        reference.cells().map(|r| r.coordinate).collect::<Vec<_>>(),
        ["B2", "C2", "D2", "B3", "C3", "D3"]
    );
}
#[test]
fn invalid_expressions_and_bounds_are_rejected() {
    for text in [
        "SUM(A1+A2)",
        "SUM(SUM(A1))",
        "SUM(A0)",
        "SUM(XFE1)",
        "SUM(A1048577)",
        "SUM(B2:A1)",
        "SUM('[book.xlsx]表'!A1)",
        "FUN_999(A1)",
        "SUM A1）",
        "WAVG(A1)",
        "COUNT(A1,B1)",
        "OTHER(A1)",
    ] {
        assert!(
            parse_rule(text, "表", "D8", &bindings(), &FunctionRegistry::default()).is_err(),
            "{text}"
        );
    }
    for text in ["普通说明", "SUM摘要", "计算金额 123"] {
        assert!(
            parse_rule(text, "表", "D8", &bindings(), &FunctionRegistry::default())
                .unwrap()
                .is_none()
        );
    }
}
#[test]
fn numeric_text_blank_percent_grouping_exponent() {
    for text in ["", " \t ", "\u{200b}\u{feff}"] {
        assert!(parse_text(text).unwrap().blank);
    }
    for (text, expected) in [
        ("0", "0"),
        ("-12.5", "-12.5"),
        ("1,234.50", "1234.5"),
        (" 3.5％ ", "0.035"),
        ("2e-3", "0.002"),
        ("\u{200b} 8%\u{feff}", "0.08"),
        ("0e999999", "0"),
        ("1e1002%", "1e1000"),
    ] {
        let value = parse_text(text).unwrap();
        assert!(!value.blank);
        assert_eq!(
            value.value,
            expected.parse::<BigDecimal>().unwrap(),
            "{text}"
        );
    }
    for text in ["NaN", "Inf", "1,23", "abc", "10%%", "1e1001", "1e-1001"] {
        assert!(parse_text(text).is_err(), "{text}");
    }
}
#[test]
fn division_has_fifty_digits_and_half_even_rounding() {
    assert_eq!(
        divide(&1.into(), &3.into()).to_string(),
        format!("0.{}", "3".repeat(50))
    );
    assert_eq!(
        divide(&2.into(), &3.into()).to_string(),
        format!("0.{}7", "6".repeat(49))
    );
    assert_eq!(
        divide(&(-10).into(), &4.into()),
        "-2.5".parse::<BigDecimal>().unwrap()
    );
    for (last, expected) in [("05", "0"), ("15", "2")] {
        let n = format!("1{}{last}", "0".repeat(48))
            .parse::<BigDecimal>()
            .unwrap();
        let result = divide(&n, &1.into());
        assert_eq!(
            result,
            format!("1{}{expected}0", "0".repeat(48))
                .parse::<BigDecimal>()
                .unwrap()
        );
    }
}
fn calculate(name: &str, rows: &[Vec<&str>]) -> excel_summary::functions::CalculationResult {
    let registry = FunctionRegistry::default();
    let mut calculator = (registry.get(name).unwrap().factory)();
    for row in rows {
        calculator.add(
            &row.iter()
                .map(|value| parse_text(value).unwrap())
                .collect::<Vec<_>>(),
        );
    }
    calculator.finish(rows.len())
}
#[test]
fn eight_algorithms_and_blank_semantics() {
    for (name, rows, expected) in [
        (
            "SUM",
            vec![vec!["1"], vec!["-2"], vec![""], vec!["0"]],
            "-1",
        ),
        ("AVG", vec![vec!["10"], vec![""]], "5"),
        ("MAX", vec![vec![""], vec!["-3"], vec!["-1"]], "-1"),
        ("MIN", vec![vec![""], vec!["3"], vec!["1"]], "1"),
        ("COUNT", vec![vec![""], vec!["0"], vec!["-1"]], "2"),
        ("SUMPRODUCT", vec![vec!["2", "3"], vec!["-4", "5"]], "-14"),
        ("RATIO", vec![vec!["1", "0"], vec!["4", "10"]], "0.5"),
        ("WAVG", vec![vec!["10", "1"], vec!["30", "3"]], "25"),
    ] {
        assert_eq!(
            calculate(name, &rows).value,
            expected.parse::<BigDecimal>().unwrap(),
            "{name}"
        );
    }
}
#[test]
fn weighted_notices_and_zero_denominators() {
    let registry = FunctionRegistry::default();
    let mut calculator = (registry.get("WAVG").unwrap().factory)();
    let notices = calculator.add(&[parse_text("").unwrap(), parse_text("-2").unwrap()]);
    assert_eq!(
        notices
            .iter()
            .map(|n| (n.code.as_str(), n.parameter_index))
            .collect::<Vec<_>>(),
        [
            ("blank_rate_with_weight", Some(0)),
            ("negative_weight", Some(1))
        ]
    );
    assert_eq!(calculator.finish(1).value, BigDecimal::zero());
    for (name, code) in [("WAVG", "zero_weight"), ("RATIO", "zero_denominator")] {
        let result = calculate(name, &[vec!["1", "1"], vec!["2", "-1"]]);
        assert!(
            result
                .notices
                .iter()
                .any(|n| n.severity == Severity::Error && n.code == code)
        );
    }
    assert!(
        (registry.get("AVG").unwrap().factory)()
            .finish(0)
            .notices
            .iter()
            .any(|notice| notice.severity == Severity::Error)
    );
}
#[test]
fn extrema_preserve_ties_ordinals_and_decimal_equality() {
    let result = calculate("MAX", &[vec![""], vec!["0"], vec!["0.000"], vec!["-1"]]);
    assert_eq!(
        result
            .sources
            .iter()
            .map(|source| source.source_index)
            .collect::<Vec<_>>(),
        [1, 2]
    );
    for name in ["MAX", "MIN"] {
        let result = calculate(name, &[vec![""], vec![""]]);
        assert!(result.sources.is_empty());
        assert!(result.notices.iter().any(|n| n.code == "no_valid_values"));
        assert!(!result.source_note.is_empty());
    }
    let result = calculate(
        "MAX",
        &[vec!["1.0000000000000000001"], vec!["1.0000000000000000002"]],
    );
    assert_eq!(result.sources[0].source_index, 1);
}
#[test]
fn arithmetic_uses_fifty_digit_context_and_fresh_state() {
    assert!(
        calculate("SUM", &[vec!["1e50"], vec!["1"], vec!["-1e50"]])
            .value
            .is_zero()
    );
    assert_eq!(calculate("SUM", &[vec!["1"]]).value, BigDecimal::from(1));
}
