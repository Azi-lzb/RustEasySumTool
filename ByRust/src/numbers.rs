use bigdecimal::num_bigint::BigInt;
use bigdecimal::num_traits::{Signed, Zero};
use bigdecimal::{BigDecimal, Context, RoundingMode};
use num_integer::Integer;
use regex::Regex;
use std::sync::OnceLock;

#[derive(Clone, Debug)]
pub struct Number {
    pub value: BigDecimal,
    pub blank: bool,
}

impl Number {
    pub fn blank() -> Self {
        Self {
            value: BigDecimal::zero(),
            blank: true,
        }
    }
    pub fn present(value: BigDecimal) -> Self {
        Self {
            value,
            blank: false,
        }
    }
}

pub fn context() -> &'static Context {
    static CONTEXT: OnceLock<Context> = OnceLock::new();
    CONTEXT.get_or_init(|| Context::new(50.try_into().unwrap(), RoundingMode::HalfEven))
}

pub fn parse_text(value: &str) -> std::result::Result<Number, String> {
    let cleaned = value.replace(['\u{200b}', '\u{feff}'], "");
    let text = cleaned.trim();
    if text.is_empty() {
        return Ok(Number::blank());
    }
    let percent = text.ends_with(['%', '％']);
    let numeric = if percent {
        text[..text.len() - text.chars().last().unwrap().len_utf8()].trim()
    } else {
        text
    };
    static PATTERN: OnceLock<Regex> = OnceLock::new();
    let pattern = PATTERN.get_or_init(|| {
        Regex::new(
        r"^[+-]?(?:(?:[0-9]{1,3}(?:,[0-9]{3})+|[0-9]+)(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$"
    ).unwrap()
    });
    if !pattern.is_match(numeric) {
        return Err(format!(
            "无法将“{}”解析成数字，非空文字不能按0处理。",
            value.chars().take(100).collect::<String>()
        ));
    }
    let mut number: BigDecimal = numeric
        .replace(',', "")
        .parse()
        .map_err(|_| "数字格式无效。")?;
    if number.is_zero() {
        return Ok(Number::present(BigDecimal::zero()));
    }
    if percent {
        let (coefficient, scale) = number.as_bigint_and_exponent();
        number = context().round_decimal(BigDecimal::new(
            coefficient,
            scale.checked_add(2).ok_or("数字的指数过大。")?,
        ));
    }
    if !number.is_zero() {
        let (_, scale) = number.as_bigint_and_exponent();
        let adjusted = i128::from(number.digits()) - i128::from(scale) - 1;
        if adjusted.abs() > 1000 {
            return Err("数字的指数过大，无法安全计算。".into());
        }
    }
    Ok(Number::present(number))
}

/// Divide with exactly 50 significant decimal digits and half-even rounding.
/// Integer division avoids dependence on BigDecimal's compile-time default precision.
pub fn divide(numerator: &BigDecimal, denominator: &BigDecimal) -> BigDecimal {
    assert!(
        !denominator.is_zero(),
        "division by zero must be validated by the calculator"
    );
    if numerator.is_zero() {
        return BigDecimal::zero();
    }
    let (n, ns) = numerator.as_bigint_and_exponent();
    let (d, ds) = denominator.as_bigint_and_exponent();
    let negative = n.sign() != d.sign();
    let n = n.abs();
    let d = d.abs();
    let difference = n.to_str_radix(10).len() as i64 - d.to_str_radix(10).len() as i64;
    let leading_less = if difference >= 0 {
        n < &d * BigInt::from(10u8).pow(difference as u32)
    } else {
        &n * BigInt::from(10u8).pow((-difference) as u32) < d
    };
    let exponent = difference - i64::from(leading_less) + ds - ns;
    let scale = 49 - exponent;
    let shift = scale + ds - ns;
    let (scaled_n, scaled_d) = if shift >= 0 {
        (n * BigInt::from(10u8).pow(shift as u32), d)
    } else {
        (n, d * BigInt::from(10u8).pow((-shift) as u32))
    };
    let (mut quotient, remainder) = scaled_n.div_rem(&scaled_d);
    let twice_remainder = remainder * 2u8;
    if twice_remainder > scaled_d || (twice_remainder == scaled_d && quotient.is_odd()) {
        quotient += 1u8;
    }
    if negative {
        quotient = -quotient;
    }
    BigDecimal::new(quotient, scale)
}

pub(crate) fn excel_number(value: &BigDecimal) -> std::result::Result<f64, String> {
    static LIMIT: OnceLock<BigDecimal> = OnceLock::new();
    let limit = LIMIT.get_or_init(|| "1.7976931348623157e308".parse().unwrap());
    if value.abs() > *limit {
        return Err("结果不是可写入Excel的有限数值。".into());
    }
    value
        .to_string()
        .parse::<f64>()
        .ok()
        .filter(|number| number.is_finite())
        .ok_or_else(|| "结果不是可写入Excel的有限数值。".into())
}
