def test_execute_rfc(sap_session):
    """Test that we can retrieve a generic RFC object."""
    rfc = sap_session.execute_rfc("RFC_READ_TABLE")
    assert rfc is not None
    assert rfc.Name == "RFC_READ_TABLE"


def test_read_table(sap_session):
    """Test that we can read a standard SAP table (T000) using the unified RFC implementation."""
    # T000 is the client table, it's very small and safe to read
    results = sap_session.read_table("T000")

    # We should get at least one row back (the client we are logged into)
    assert len(results) > 0
    assert isinstance(results, list)
    assert isinstance(results[0], dict)

    # Let's read with specific fields and options
    fields = ["MANDT", "MTEXT"]
    options = ["MANDT = '900'"]  # Testing against our test client if possible

    filtered_results = sap_session.read_table("T000", fields=fields, options=options)
    print(filtered_results)
    assert isinstance(filtered_results, list)
