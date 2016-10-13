require 'omniauth-oauth2'

module OmniAuth
  module Strategies
    class Office365 < OmniAuth::Strategies::OAuth2
      option :name, :office365
      
      DEFAULT_SCOPE="openid email profile https://outlook.office.com/contacts.read"

      option :client_options, {
        site:          'https://login.microsoftonline.com/',
        token_url:     'https://login.microsoftonline.com/415abcba-efe6-43bc-b3e7-1f1e9fb2dd3b/oauth2/v2.0/token',
        authorize_url: 'https://login.microsoftonline.com/415abcba-efe6-43bc-b3e7-1f1e9fb2dd3b/oauth2/v2.0/authorize'
      }
      
      option :authorize_options, [:scope]

      uid { raw_info["objectId"] }

      info do
        {
          'email' => raw_info["userPrincipalName"],
          'name' => [raw_info["givenName"], raw_info["surname"]].join(' '),
          'nickname' => raw_info["displayName"]
        }
      end

      extra do
        {
          'raw_info' => raw_info
        }
      end

      def raw_info
        @raw_info ||= access_token.get("https://outlook.office.com/api/v2.0/me/").parsed
      end
    end
  end
end
