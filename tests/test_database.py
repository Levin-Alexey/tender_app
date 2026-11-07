import unittest

from database import normalize_database_url


class NormalizeDatabaseUrlTests(unittest.TestCase):
    def test_normalizes_asyncpg_sqlalchemy_scheme(self):
        self.assertEqual(
            normalize_database_url(
                "postgresql+asyncpg://user:password@db.example/tenderdb"
            ),
            "postgresql://user:password@db.example/tenderdb",
        )

    def test_keeps_native_postgres_scheme(self):
        database_url = "postgresql://user:password@db.example/tenderdb"
        self.assertEqual(normalize_database_url(database_url), database_url)

    def test_rejects_missing_or_unsupported_urls(self):
        for database_url in ("", "mysql://db.example/tenderdb"):
            with self.subTest(database_url=database_url):
                with self.assertRaises(RuntimeError):
                    normalize_database_url(database_url)


if __name__ == "__main__":
    unittest.main()
