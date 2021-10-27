require_relative '../spec_helper'

describe Viewpoint::EWS::ConvertAccessors do
  let(:ews) do
    con = double('Connection')
    Viewpoint::EWS::SOAP::ExchangeWebService.new con,
      {:server_version => Viewpoint::EWS::SOAP::VERSION_2010_SP2}
  end

  let(:client) do
    client = double("EWSClient")
    client.extend subject
    client.stub(:ews) {ews}
    client
  end

  before do
    ews.stub(:do_soap_request)
  end

  it "generates ConvertId XML" do
    ews.should_receive(:do_soap_request).
      with(match_xml(load_soap("convert_id", :request)))

    id = "AAMkAGQyMmE3ODAxLWZlNTItNDdiZS04NzNhLWQxNWQ4MjUzMzY1YgAuAAAAAADiDiwt8lJHQI9Vqfh/HOkiAQDwNDPGgnbYTa3P6qFapoP/AAAAAAENAAA="

    opts = {
      format: :ews_id,
      destination_format: :entry_id,
      mailbox: 'foo@bar.com',
    }
    client.convert_id(id, opts)
  end
end
