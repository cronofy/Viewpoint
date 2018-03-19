require_relative '../spec_helper'

describe "Tolerant parsing" do
  it 'parses a successful response' do
    soap_resp = load_soap 'dodgy_ews', :response
    resp = Viewpoint::EWS::SOAP::EwsParser.new(soap_resp).parse(response_class: Viewpoint::EWS::SOAP::EwsResponse)
    expect { resp.body }.to_not raise_error
  end

  it 'raises a specific error when parsing empry responses' do
    soap_resp = ""
    parser = Viewpoint::EWS::SOAP::EwsParser.new(soap_resp)
    expect { parser.parse(response_class: Viewpoint::EWS::SOAP::EwsResponse) }.to raise_error(RuntimeError)
  end

  it 'raises a specific error when valid xml but invalid ews responses' do
    soap_resp = "<xml></xml>"
    parser = Viewpoint::EWS::SOAP::EwsParser.new(soap_resp)
    expect { parser.parse(response_class: Viewpoint::EWS::SOAP::EwsResponse) }.to raise_error(Viewpoint::EWS::Errors::MalformedResponseError)
  end
end
