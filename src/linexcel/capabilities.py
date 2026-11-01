"""Versioned capability declarations; never a certification of Excel parity."""

from importlib.metadata import version


def capability_registry() -> dict:
    verified = "independent_regressions_only"
    return {
        "schemaVersion": 2,
        "engine": {"name": "formualizer", "version": version("formualizer")},
        "verification": verified,
        "nativeExcelCertification": "not_established",
        "features": [
            {
                "id": "structured_tables",
                "extraction": "declared_ooxml_metadata",
                "resolution": "static_source_rectangles",
                "calculation": "native_engine_with_resolved_table_references",
                "verification": verified,
                "limitations": [
                    "empty data regions",
                    "disjoint selections",
                    "invalid or ambiguous table metadata",
                ],
            },
            {
                "id": "defined_names",
                "extraction": "declared_ooxml_metadata",
                "resolution": "scoped_constants_absolute_refs_static_computed_names",
                "calculation": "bounded_native_expression_expansion",
                "verification": verified,
                "limitations": [
                    "relative references without an established definition anchor",
                    "dynamic or cyclic definitions",
                    "expansion over 32 names or 8192 characters",
                ],
            },
            {
                "id": "three_dimensional_references",
                "extraction": "source_sheet_order_and_formula_tokens",
                "resolution": "exact_sheet_span",
                "calculation": "documented_direct_aggregate_arguments",
                "verification": verified,
                "limitations": [
                    "non-aggregate contexts",
                    "computed reference arguments",
                ],
            },
            {
                "id": "external_references",
                "extraction": "declared_links_and_explicitly_supplied_files",
                "resolution": "transitive_sparse_cells_ranges_names_tables",
                "calculation": "native_formulas_without_cache_substitution",
                "verification": verified,
                "limitations": [
                    "missing or unprovided workbooks",
                    "mixed-calendar formulas and dates",
                    "external array regions",
                    "external dynamic reference context",
                ],
            },
            {
                "id": "visibility_aggregates",
                "extraction": "source_row_visibility_table_and_sheet_filters",
                "resolution": "bounded_reference_form_masks",
                "calculation": "SUBTOTAL_1_to_11_101_to_111_AGGREGATE_1_to_19",
                "verification": verified,
                "limitations": [
                    "active filters with modes including manually hidden rows",
                    "computed reference arguments or nonliteral options",
                    "external visibility context",
                    "more than 20000 reference positions",
                    "nonliteral or fractional rank/quartile arguments",
                    "empty MEDIAN requires native Excel verification",
                    "tied MODE.SNGL results require native Excel verification",
                ],
            },
            {
                "id": "arrays_and_spills",
                "extraction": "declared_ooxml_array_regions",
                "resolution": "anchor_dependencies_including_unsaved_spill_children",
                "calculation": "single_native_batch_with_extent_verification",
                "verification": verified,
                "limitations": [
                    "undeclared or changing spill extents",
                    "data tables",
                    "external array regions",
                    "cyclic array regions",
                ],
            },
            {
                "id": "dynamic_references",
                "extraction": "original_formula_and_checked_source_expressions",
                "resolution": "INDIRECT_A1_R1C1_OFFSET_contiguous_rectangles",
                "calculation": "native_formula_with_exact_resolved_closure",
                "verification": verified,
                "limitations": [
                    "address expressions outside bounded scalar arithmetic "
                    "and concatenation",
                    "runtime or volatile address expressions",
                    "dynamic defined names",
                    "fractional OFFSET arguments",
                    "external dynamic contexts",
                ],
            },
            {
                "id": "volatile_context",
                "extraction": "formula_tokens",
                "resolution": "operation_clock_UTC_and_uint64_seed",
                "calculation": "replayable_native_operation_context",
                "verification": verified,
                "limitations": ["Excel random sequence parity is not established"],
            },
            {
                "id": "iterative_cycles",
                "extraction": "source_iteration_settings",
                "resolution": "exact_dependency_closure",
                "calculation": "native_iteration_with_convergence_telemetry",
                "verification": verified,
                "limitations": [
                    "disabled or malformed settings",
                    "nonconvergence",
                    "different settings across external cyclic sources",
                    "saved iteration initial states are not substituted",
                ],
            },
            {
                "id": "other_workbook_sensitive_functions",
                "extraction": "formula_tokens",
                "resolution": "unsupported",
                "calculation": "refused",
                "verification": "guard_regressions",
                "limitations": [
                    "CELL/INFO/SHEET/SHEETS/FORMULATEXT",
                    "RTD/WEBSERVICE/STOCKHISTORY",
                ],
            },
        ],
    }
