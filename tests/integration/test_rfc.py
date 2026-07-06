def test_execute_rfc(sap_session):
    """Test that we can retrieve a generic RFC object."""
    rfc = sap_session.rfc.execute_rfc("RFC_READ_TABLE")
    assert rfc is not None
    assert rfc.Name == "RFC_READ_TABLE"


def test_read_table_active_session(sap_session):
    """Test that we can read a standard SAP table (T000) via active session's rfc property."""
    # T000 is the client table, it's very small and safe to read
    results = sap_session.rfc.read_table("T000")

    # We should get at least one row back (the client we are logged into)
    assert len(results) > 0
    assert isinstance(results, list)
    assert isinstance(results[0], dict)

    # Let's read with specific fields and options
    fields = ["MANDT", "MTEXT"]
    options = ["MANDT = '900'"]  # Testing against our test client if possible

    filtered_results = sap_session.rfc.read_table("T000", fields=fields, options=options)
    print(filtered_results)
    assert isinstance(filtered_results, list)

def test_read_table_headless(rfc_session):
    """Test that we can read a standard SAP table (T000) using the unified RFC implementation."""
    # T000 is the client table, it's very small and safe to read
    results = rfc_session.read_table("T000")

    # We should get at least one row back (the client we are logged into)
    assert len(results) > 0
    assert isinstance(results, list)
    assert isinstance(results[0], dict)

    # Let's read with specific fields and options
    fields = ["MANDT", "MTEXT"]
    options = ["MANDT = '900'"]  # Testing against our test client if possible

    filtered_results = rfc_session.read_table("T000", fields=fields, options=options)
    print(filtered_results)
    assert isinstance(filtered_results, list)

def test_call_bapi_headless(rfc_session, secrets):
    """Test executing a BAPI using the headless RFC connection."""
    user_id = secrets.get("user_id", "SAP*").upper()
    
    result = rfc_session.call_bapi(
        bapi_name="BAPI_USER_GET_DETAIL",
        import_params={"USERNAME": user_id},
        extract_tables=["RETURN", "ACTIVITYGROUPS"],
        extract_imports=["ADDRESS"]
    )
    
    assert "EXPORTS" in result
    assert "ADDRESS" in result["EXPORTS"]
    assert "TABLES" in result
    assert "RETURN" in result["TABLES"]

def test_call_bapi_active_session(sap_session, secrets):
    """Test executing a BAPI using the active session's rfc property."""
    user_id = secrets.get("user_id", "SAP*").upper()
    
    result = sap_session.rfc.call_bapi(
        bapi_name="BAPI_USER_GET_DETAIL",
        import_params={"USERNAME": user_id},
        extract_tables=["RETURN", "ACTIVITYGROUPS"],
        extract_imports=["ADDRESS"]
    )
    
    assert "EXPORTS" in result
    assert "ADDRESS" in result["EXPORTS"]
    assert "TABLES" in result
    assert "RETURN" in result["TABLES"]