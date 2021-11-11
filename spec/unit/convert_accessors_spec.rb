require_relative '../spec_helper'

describe Viewpoint::EWS::ConvertAccessors do
  let(:ews) do
    con = double('Connection')
    Viewpoint::EWS::SOAP::ExchangeWebService.new con,
      {:server_version => Viewpoint::EWS::SOAP::VERSION_2013}
  end

  let(:client) do
    client = double("EWSClient")
    client.extend described_class
    client.stub(:ews) {ews}
    client
  end

  let(:response) do
    Viewpoint::EWS::SOAP::EwsParser.new(load_soap("convert_id", :response))
      .parse(response_class: Viewpoint::EWS::SOAP::EwsResponse)
  end

  subject do
    id = "AAMkAGQyMmE3ODAxLWZlNTItNDdiZS04NzNhLWQxNWQ4MjUzMzY1YgAuAAAAAADiDiwt8lJHQI9Vqfh/HOkiAQDwNDPGgnbYTa3P6qFapoP/AAAAAAENAAA="

    opts = {
      format: :ews_id,
      destination_format: :entry_id,
      mailbox: 'foo@bar.com',
    }

    client.convert_id(id, opts)
  end

  it "generates ConvertId XML" do
    ews.should_receive(:do_soap_request).
      with(match_xml(load_soap("convert_id", :request)), response_class: Viewpoint::EWS::SOAP::EwsResponse)
      .and_return(response)

    subject
  end

  it "returns the converted ID" do
    ews.stub(:do_soap_request).and_return(response)

    expect(subject).to eq('AAAAAOIOLC3yUkdAj1Wp+H8c6SIBAPA0M8aCdthNrc/qoVqmg/8AAAAAAQ0AAA==')
  end
end
