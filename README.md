# Omniauth::Office365 [![Build Status](https://travis-ci.org/simi/omniauth-office365.svg?branch=master)](https://travis-ci.org/simi/omniauth-office365)

Office365 OAuth2 Strategy for OmniAuth. This gem is for anyone looking to use Office 365's v2.0 api oauth2.

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'omniauth-office365-v2.0'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install omniauth-office365-v2.0

## Usage

```ruby
Rails.application.config.middleware.use OmniAuth::Builder do
  provider :office365, ENV['OFFICE365_KEY'], ENV['OFFICE365_SECRET']
end
```


## Contributing

1. Fork it ( https://github.com/simi/omniauth-office365/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request
