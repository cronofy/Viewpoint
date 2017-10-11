require_relative '../spec_helper'

describe "Operations on Exchange Data Services" do

  before do
    con = double('Connection')
    @ews = Viewpoint::EWS::SOAP::ExchangeWebService.new con,
      {:server_version => Viewpoint::EWS::SOAP::VERSION_2010_SP2}
    @ews.stub(:do_soap_request)
  end

  it "generates CreateItem for calendar item XML" do
    expected_doc = load_soap("create_item", :request)

    @ews.should_receive(:do_soap_request) do |request_document|
      request_document.to_xml.should eq(Nokogiri.XML(expected_doc).to_xml)
    end

    opts = {
      send_meeting_invitations: 'SendToAllAndSaveCopy',
      items: [
        {
          calendar_item: {
            subject: 'test cal item',
            body: {body_type: 'Text', text: 'this is a test cal item'},
            start: {text: "2017-10-11T10:11:08+01:000"},
            end: {text: "2017-10-11T11:11:08+01:000"},
            required_attendees: [
              { attendee: { mailbox: { email_address: 'example@example.com'}}},
            ],
            enhanced_location: {
              display_name:  "Easy street london",
              annotation: "The cavern",
              postal_address: {
                street: "Easy street",
                city: "London",
                state: "Londonshire",
                country: "United Kingdom",
                postal_code: "A11AA",
                type: "Business",
                latitude: "50.0222",
                longitude: "3.1415",
                accuracy: "100",
                altitude: "50",
                altitude_accuracy: "100",
                formatted_address: "Easy street london",
                location_uri: "http://www.example.com",
                location_source: "Device",
              }
            }
          }
        }
      ]
    }

    @ews.create_item(opts)
  end
end
