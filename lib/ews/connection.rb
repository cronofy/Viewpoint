=begin
  This file is part of Viewpoint; the Ruby library for Microsoft Exchange Web Services.

  Copyright © 2011 Dan Wanek <dan.wanek@gmail.com>

  Licensed under the Apache License, Version 2.0 (the "License");
  you may not use this file except in compliance with the License.
  You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

  Unless required by applicable law or agreed to in writing, software
  distributed under the License is distributed on an "AS IS" BASIS,
  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
  See the License for the specific language governing permissions and
  limitations under the License.
=end
require 'httpclient'

class Viewpoint::EWS::Connection
  include Viewpoint::EWS::ConnectionHelper
  include Viewpoint::EWS

  attr_reader :endpoint
  # @param [String] endpoint the URL of the web service.
  #   @example https://<site>/ews/Exchange.asmx
  # @param [Hash] opts Misc config options (mostly for developement)
  # @option opts [Fixnum] :ssl_verify_mode
  # @option opts [Fixnum] :receive_timeout override the default receive timeout
  #   seconds
  # @option opts [Fixnum] :connect_timeout override the default connect timeout
  #   seconds
  # @option opts [Array]  :trust_ca an array of hashed dir paths or a file
  # @option opts [String] :user_agent the http user agent to use in all requests
  def initialize(endpoint, opts = {})
    @log = Logging.logger[self.class.name.to_s.to_sym]
    @httpcli = HTTPClient.new

    if opts[:trust_ca]
      @httpcli.ssl_config.clear_cert_store
      opts[:trust_ca].each do |ca|
        @httpcli.ssl_config.add_trust_ca ca
      end
    end

    if opts[:debug_dev]
      @httpcli.debug_dev = opts[:debug_dev]
    end

    if opts[:force_basic_auth]
      @httpcli.force_basic_auth = true
    end

    @httpcli.default_header['User-Agent'] = opts[:user_agent] if opts[:user_agent]

    @httpcli.ssl_config.verify_mode = opts[:ssl_verify_mode] if opts[:ssl_verify_mode]
    @httpcli.ssl_config.ssl_version = opts[:ssl_version] if opts[:ssl_version]
    @httpcli.ssl_config.options |= OpenSSL::SSL::OP_LEGACY_SERVER_CONNECT if opts[:ssl_legacy_server_connect]
    # Up the keep-alive so we don't have to do the NTLM dance as often.
    @httpcli.keep_alive_timeout = 60
    @httpcli.connect_timeout = opts[:connect_timeout] if opts[:connect_timeout]
    @httpcli.send_timeout = opts[:send_timeout] if opts[:send_timeout]
    @httpcli.receive_timeout = opts[:receive_timeout] if opts[:receive_timeout]
    @httpcli.connect_timeout = opts[:connect_timeout] if opts[:connect_timeout]
    @httpcli.transparent_gzip_decompression = true
    @endpoint = endpoint
  end

  def force_basic_auth=(value)
    @httpcli.force_basic_auth = value
  end

  def set_auth(user,pass)
    @httpcli.default_header.delete('Authorization')
    @httpcli.set_auth(@endpoint.to_s, user, pass)
  end

  def set_oauth_access_token(access_token)
    @httpcli.default_header['Authorization'] = "Bearer #{access_token}"
  end

  # Authenticate to the web service. You don't have to do this because
  # authentication will happen on the first request if you don't do it here.
  # @return [Boolean] true if authentication is successful, false otherwise
  def authenticate
    self.get && true
  end

  # Every Connection class must have the dispatch method. It is what sends the
  # SOAP request to the server and calls the parser method on the EWS instance.
  #
  # This was originally in the ExchangeWebService class but it was added here
  # to make the processing chain easier to modify. For example, it allows the
  # reactor pattern to handle the request with a callback.
  # @param ews [Viewpoint::EWS::SOAP::ExchangeWebService] used to call
  #   #parse_soap_response
  # @param soapmsg [String]
  # @param opts [Hash] misc opts for handling the Response
  def dispatch(ews, soapmsg, opts)
    if log.debug?
      log.debug { "REQUEST =============" }
      log.debug soapmsg
    end

    headers = {}

    set_mailbox_headers(ews, headers)

    respmsg = post(soapmsg, headers)
    @log.debug <<-EOF.gsub(/^ {6}/, '')
      Received SOAP Response:
      ----------------
      #{Nokogiri::XML(respmsg).to_xml}
      ----------------
    EOF

    if log.debug?
      log.debug { "RESPONSE =============" }
      log.debug Nokogiri::XML(respmsg).to_xml
    end

    opts[:raw_response] ? respmsg : ews.parse_soap_response(respmsg, opts)
  end

  def set_mailbox_headers(ews, headers)
    return unless present?(ews.anchor_mailbox)

    headers['X-AnchorMailbox'] = ews.anchor_mailbox
    headers['X-PreferServerAffinity'] = 'true'
    log.debug { "HEADERS =============" }
    log.debug headers
  end

  def present?(string)
    string && !string.empty?
  end

  # Send a GET to the web service
  # @return [String] If the request is successful (200) it returns the body of
  #   the response.
  def get
    check_response( @httpcli.get(@endpoint) )
  end

  # Send a POST to the web service
  # @return [String] If the request is successful (200) it returns the body of
  #   the response.
  def post(xmldoc, additional_headers = {})
    headers = {'Content-Type' => 'text/xml', 'Accept-Encoding' => 'gzip;q=1.0,deflate;q=0.6,identity;q=0.3' }
    check_response( @httpcli.post(@endpoint, xmldoc, headers.merge(additional_headers)) )
  end

  private

  def check_response(resp)
    case resp.status
    when 200
      resp.body
    when 302
      # @todo redirect
      raise Errors::UnhandledResponseError.new("Unhandled HTTP Redirect", resp)
    when 401
      raise Errors::UnauthorizedResponseError.new("Unauthorized request", resp)
    when 403
      raise Errors::ForbiddenResponseError.new("Forbidden request", resp)
    when 500
      if resp.headers['Content-Type'] =~ /xml/
        err_string, err_code = parse_soap_error(resp.body)
        raise Errors::SoapResponseError.new("SOAP Error: Message: #{err_string}  Code: #{err_code} - #{resp.body}", resp, err_code, err_string)
      else
        raise Errors::ServerError.new("Internal Server Error. Message: #{resp.body}", resp)
      end
    when 510
      raise Errors::NotExtendedResponseError.new("The policy for accessing the resource has not been met in the request.", resp)
    else
      raise Errors::ResponseError.new("HTTP Error Code: #{resp.status}, Msg: #{resp.body}", resp)
    end
  end

  # @param [String] xml to parse the errors from.
  def parse_soap_error(xml)
    ndoc = Nokogiri::XML(xml)
    ns = ndoc.collect_namespaces
    err_string  = ndoc.xpath("//faultstring",ns).text
    err_code    = ndoc.xpath("//faultcode",ns).text
    @log.debug "Internal SOAP error. Message: #{err_string}, Code: #{err_code}"
    [err_string, err_code]
  end

end
