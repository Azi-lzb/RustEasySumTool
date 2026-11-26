//! Excel template aggregation with calamine reads and umya-spreadsheet writes.
//!
//! No Python runtime is required. A run snapshots `config.xlsx`, validates every
//! template instruction, accumulates source values using 50 decimal digits, then
//! publishes a new workbook without replacing an existing file. Original business
//! XML is retained around changed cells to protect settings lost by umya round trips.
//!
//! ```no_run
//! use excel_summary::{aggregate, AggregationRequest, RunControl};
//! let request = AggregationRequest {
//!     template_path: "template.xlsx".into(),
//!     source_paths: vec!["source-a.xlsx".into(), "source-b.xlsx".into()],
//!     output_path: "result.xlsx".into(),
//!     config_path: Some("config.xlsx".into()),
//! };
//! let result = aggregate(&request, &RunControl::default())?;
//! println!("{}", result.output_path.display());
//! # Ok::<(), excel_summary::AggregationError>(())
//! ```
//!
//! [`Aggregator`] exposes a [`functions::FunctionRegistry`] for application-specific
//! calculators. [`RunControl`] accepts a progress callback and an atomic cancel flag.
//! [`model::Issue`] keeps source and template locations for UI display.
pub mod configuration;
pub mod engine;
pub mod functions;
pub mod model;
pub mod numbers;
mod package;
mod preserve;
pub mod provenance;
pub mod reader;
mod reports;
pub mod rules;
pub use engine::{Aggregator, aggregate, default_output_path};
pub use model::{AggregationError, AggregationRequest, AggregationResult, RunControl};
