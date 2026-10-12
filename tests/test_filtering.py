import pytest

from ankiros_sync import fetch_exhibitors, filter_exhibitors


RECORDS = [
    {
        "company_name": "Direct A",
        "relationship": "direct",
        "parent_company": None,
    },
    {
        "company_name": "Co A",
        "relationship": "co-exhibitor",
        "parent_company": "Direct A",
    },
]


def test_filter_returns_direct_exhibitors_only():
    assert [r["company_name"] for r in filter_exhibitors(RECORDS, "direct")] == [
        "Direct A"
    ]


def test_filter_returns_co_exhibitors_only():
    assert [
        r["company_name"]
        for r in filter_exhibitors(RECORDS, "co-exhibitors")
    ] == ["Co A"]


def test_filter_all_preserves_both_relationships():
    assert filter_exhibitors(RECORDS, "all") == RECORDS


def test_filter_rejects_unknown_value():
    with pytest.raises(ValueError):
        filter_exhibitors(RECORDS, "partners")


def test_fetch_expands_subco_companies():
    class Session:
        def post(self, url, **kwargs):
            class Response:
                status_code = 200
                headers = {}

                def raise_for_status(self):
                    return None

                def json(self):
                    if url.endswith("getAllParticipantList"):
                        return {
                            "result": [
                                {"companyName": "Direct A", "slug": "direct-a"}
                            ]
                        }
                    return {
                        "result": {
                            "subcoCompany": [
                                {"companyName": "Co A", "slug": "co-a"}
                            ]
                        }
                    }

            return Response()

    records = fetch_exhibitors(Session())
    assert records == [
        {
            "company_name": "Direct A",
            "slug": "direct-a",
            "relationship": "direct",
            "parent_company": None,
        },
        {
            "company_name": "Co A",
            "slug": "co-a",
            "relationship": "co-exhibitor",
            "parent_company": "Direct A",
        },
    ]
