def is_pass(val, nominal, plus, minus):
    """
    Returns True if val is within [nominal - minus - 0.005, nominal + plus + 0.005], else False.
    Validates using 3 decimal places with rounding margin.
    
    Args:
        val: measured value (float or convertible to float)
        nominal: nominal/reference value (float or None)
        plus: upper tolerance (float or convertible)
        minus: lower tolerance (float or convertible)
    
    Returns:
        True if val is within tolerance range, False otherwise.
        Returns False if nominal is None or any conversion fails.
    """
    try:
        # Round to 3 decimals for validation
        val_rounded = round(float(val), 3)
        nom_float = float(nominal) if nominal is not None else None
        plus_float = float(plus) if plus is not None else 0.0
        minus_float = float(minus) if minus is not None else 0.0
        
        if nom_float is None:
            return False
        
        # Bounds with 0.005 margin for 3-decimal rounding tolerance
        low = round(nom_float - minus_float - 0.005, 3)
        high = round(nom_float + plus_float + 0.005, 3)
        
        return low <= val_rounded <= high
    
    except (ValueError, TypeError):
        return False


def strip_unit_symbols(name):
    """
    Remove unit symbols from column name for tolerance lookup.
    E.g., 'Diameter 1 (mm)' -> 'Diameter 1'
    """
    return name.replace(" (mm)", "").replace(" (⟳)", "").replace(" (°)", "").strip()


def validate_measurements(rows, col_names, tolerance_dict):
    """
    Validates all rows against column-specific tolerances.
    
    Args:
        rows: list of data tuple rows (parsed)
        col_names: list of ordered column names (may have units)
        tolerance_dict: {col_name_no_unit: (nom, plus, minus)}
    
    Returns:
        dict: {col_name: ["PASS"/"FAIL"/"", ...]} per row
    """
    results = {col: [] for col in col_names}
    
    for row in rows:
        for idx, col in enumerate(col_names):
            # Column index is valid in the row
            if idx >= len(row):
                results[col].append("")
                continue
            
            val = row[idx]
            col_clean = strip_unit_symbols(col)
            
            # Only check if tolerance is set and value is not empty/"-"
            if col_clean in tolerance_dict and val not in (None, "", "-"):
                try:
                    val_float = float(val)
                    nominal, plus, minus = tolerance_dict[col_clean]
                    nominal_float = float(nominal) if nominal is not None else None
                    plus_float = float(plus) if plus is not None else 0.0
                    minus_float = float(minus) if minus is not None else 0.0
                    
                    if nominal_float is None:
                        results[col].append("")
                    else:
                        status = "PASS" if is_pass(val_float, nominal_float, plus_float, minus_float) else "FAIL"
                        results[col].append(status)
                
                except (ValueError, TypeError):
                    results[col].append("")
            else:
                results[col].append("")
    
    return results


def validate_measurements_legacy(rows, headers, tolerance_dict):
    """
    Legacy: Use when data has 'Type'/'Value' columns (for classic/old formats).
    
    Args:
        rows: list of tuples (parsed)
        headers: header tuple (e.g., ('ID', 'Type', 'Value', ...))
        tolerance_dict: {'Diameter': (nominal, plus, minus), ...}
    
    Returns:
        list of dicts [{"dimension":..., "value":..., "status":..., "threshold": (...)}]
    """
    results = []
    
    if "Type" not in headers or "Value" not in headers:
        return results
    
    type_idx = headers.index("Type")
    value_idx = headers.index("Value")
    
    for row in rows:
        try:
            dim = row[type_idx]
            val = float(row[value_idx])
            
            if dim in tolerance_dict:
                nominal, plus, minus = tolerance_dict[dim]
                
                if nominal is not None:
                    status = "PASS" if is_pass(val, nominal, plus, minus) else "FAIL"
                    results.append({
                        "dimension": dim,
                        "value": val,
                        "status": status,
                        "threshold": (nominal - minus, nominal + plus)
                    })
        
        except Exception:
            continue
    
    return results
