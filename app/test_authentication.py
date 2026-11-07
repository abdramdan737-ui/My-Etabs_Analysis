import unittest

from app.authentication import auth_configuration_issue


class AuthConfigurationTests(unittest.TestCase):
    def test_missing_auth_configuration_is_rejected(self):
        self.assertIn("لم يُعثر", auth_configuration_issue(None))

    def test_missing_shared_setting_is_reported(self):
        issue = auth_configuration_issue(
            {
                "redirect_uri": "https://example.test/oauth2callback",
                "auth0": {
                    "client_id": "client",
                    "client_secret": "secret",
                    "server_metadata_url": "https://example.test/.well-known/openid-configuration",
                },
            }
        )
        self.assertIn("cookie_secret", issue)

    def test_missing_auth0_setting_is_reported(self):
        issue = auth_configuration_issue(
            {
                "redirect_uri": "https://example.test/oauth2callback",
                "cookie_secret": "random-cookie-secret",
                "auth0": {
                    "client_id": "client",
                    "server_metadata_url": "https://example.test/.well-known/openid-configuration",
                },
            }
        )
        self.assertIn("client_secret", issue)

    def test_example_placeholders_are_not_treated_as_credentials(self):
        issue = auth_configuration_issue(
            {
                "redirect_uri": "https://YOUR-APP-HOST/oauth2callback",
                "cookie_secret": "REPLACE_WITH_A_LONG_RANDOM_SECRET",
                "auth0": {
                    "client_id": "YOUR_AUTH0_CLIENT_ID",
                    "client_secret": "YOUR_AUTH0_CLIENT_SECRET",
                    "server_metadata_url": "https://YOUR_AUTH0_DOMAIN/.well-known/openid-configuration",
                },
            }
        )
        self.assertIn("redirect_uri", issue)

    def test_complete_auth0_configuration_is_accepted(self):
        issue = auth_configuration_issue(
            {
                "redirect_uri": "https://example.test/oauth2callback",
                "cookie_secret": "random-cookie-secret",
                "auth0": {
                    "client_id": "client",
                    "client_secret": "secret",
                    "server_metadata_url": "https://example.test/.well-known/openid-configuration",
                },
            }
        )
        self.assertIsNone(issue)


if __name__ == "__main__":
    unittest.main()
