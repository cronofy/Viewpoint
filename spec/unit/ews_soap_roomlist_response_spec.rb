require_relative '../spec_helper'

describe Viewpoint::EWS::SOAP::EwsSoapRoomlistResponse do

  let(:ews) { Viewpoint::EWS::SOAP::ExchangeWebService.new(double(:connection)) }

  ERROR_NAME_RESOLUTION = <<EOS.gsub(%r{>\s+}, '>')
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Header>
    <h:ServerVersionInfo xmlns:h="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" MajorVersion="15" MinorVersion="20" MajorBuildNumber="56" MinorBuildNumber="13" Version="V2017_07_11"/>
  </s:Header>
  <s:Body>
    <m:GetRoomListsResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ResponseClass="Error">
      <m:MessageText>User must have a mailbox for name resolution operations.</m:MessageText>
      <m:ResponseCode>ErrorNameResolutionNoMailbox</m:ResponseCode>
      <m:DescriptiveLinkKey>0</m:DescriptiveLinkKey>
    </m:GetRoomListsResponse>
  </s:Body>
</s:Envelope>
EOS

  context "Exchange responds with ErrorNameResolutionNoMailbox error" do
    let(:ews)  { Viewpoint::EWS::SOAP::ExchangeWebService.new(double(:connection)) }
    let(:resp) { ews.parse_soap_response(ERROR_NAME_RESOLUTION, :response_class => described_class) }

    it "should deliver the error code" do
      resp.code.should eq "ErrorNameResolutionNoMailbox"
    end

    it "should report an error" do
      resp.success?.should eq false
    end

    it "should extract the response message" do
      extracted_response_message = {
        attribs: {response_class: "Error"},
        elems: [
          {message_text:         {text: "User must have a mailbox for name resolution operations."}},
          {response_code:        {text: "ErrorNameResolutionNoMailbox"}},
          {descriptive_link_key: {text: "0"}},
        ]
      }
      resp.response_message.should == extracted_response_message
    end
  end

  SUCCESSFUL_LOOKUP = <<EOS.gsub(%r{>\s+}, '>')
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Header>
    <h:ServerVersionInfo MajorVersion="14" MinorVersion="1" MajorBuildNumber="164" MinorBuildNumber="0" Version="Exchange2010_SP1" xmlns:h="http://schemas.microsoft.com/exchange/services/2006/types" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema"/>
  </s:Header>
  <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <GetRoomListsResponse ResponseClass="Success" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
      <ResponseCode>NoError</ResponseCode>
      <m:RoomLists xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
        <t:Address xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
          <t:Name>Room List</t:Name>
          <t:EmailAddress>RoomList@contoso.com</t:EmailAddress>
          <t:RoutingType>SMTP</t:RoutingType>
          <t:MailboxType>PublicDL</t:MailboxType>
        </t:Address>
      </m:RoomLists>
    </GetRoomListsResponse>
  </s:Body>
</s:Envelope>
EOS


  context "Successful lookup" do
    let(:resp) { ews.parse_soap_response(SUCCESSFUL_LOOKUP, :response_class => described_class) }

    it "response should be successful" do
      resp.status.should eq "Success"
    end

    it "should not report an error" do
      resp.success?.should eq true
    end

    it "the roomlist should have one element" do
      expect(resp.roomListsArray).to eq(
        [
          {:address=>
           {:elems=>
            {:name=>{:text=>"Room List"},
             :email_address=>{:text=>"RoomList@contoso.com"},
             :routing_type=>{:text=>"SMTP"},
             :mailbox_type=>{:text=>"PublicDL"}}}}
        ]
      )
    end
  end
end
