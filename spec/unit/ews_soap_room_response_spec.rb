require_relative '../spec_helper'

describe Viewpoint::EWS::SOAP::EwsSoapRoomResponse do

  let(:ews) { Viewpoint::EWS::SOAP::ExchangeWebService.new(double(:connection)) }

  ROOM_SUCCESSFUL_LOOKUP = <<EOS.gsub(%r{>\s+}, '>')
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Header>
    <h:ServerVersionInfo MajorVersion="15" MinorVersion="0" MajorBuildNumber="873" MinorBuildNumber="9"
                         Version="V2_9" xmlns:h="http://scemas.microsoft.com/exchange/services/2006/types"
                         xmlns="http://schemas.microsoft.com/exchange/services/2006/types"
                         xmlns:xsd="http://www.w3org/2001/XMLSchema"
                         xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" />
  </s:Header>
  <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <m:GetRoomsResponse ResponseClass="Success" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                        xmlns:t="http://scemas.microsoft.com/exchange/services/2006/types">
      <m:ResponseCode>NoError</m:ResponseCode>
      <m:Rooms>
        <t:Room>
          <t:Id>
            <t:Name>Conf Room 3/101 (16) AV</t:Name>
            <t:EmailAddress>cf3101@contoso.com</t:EmailAddress>
            <t:RoutingType>SMTP</t:RoutingType>
            <t:MailboxType>Mailbox</t:MailboxType>
          </t:Id>
        </t:Room>
        <t:Room>
          <t:Id>
            <t:Name>Conf Room 3/102 (8) AV</t:Name>
            <t:EmailAddress>cf3102@contoso.com</t:EmailAddress>
            <t:RoutingType>SMTP</t:RoutingType>
            <t:MailboxType>Mailbox</t:MailboxType>
          </t:Id>
        </t:Room>
        <t:Room>
          <t:Id>
            <t:Name>Conf Room 3/103 (14) AV RoundTable</t:Name>
            <t:EmailAddress>cf3103@contoso.com</t:EmailAddress>
            <t:RoutingType>SMTP</t:RoutingType>
            <t:MailboxType>Mailbox</t:MailboxType>
          </t:Id>
        </t:Room>
      </m:Rooms>
    </m:GetRoomsResponse>
  </s:Body>
</s:Envelope>
EOS


  context "Successful lookup" do
    let(:resp) { ews.parse_soap_response(ROOM_SUCCESSFUL_LOOKUP, :response_class => described_class) }

    it "response should be successful" do
      resp.status.should eq "Success"
    end

    it "should not report an error" do
      resp.success?.should eq true
    end

    it "the room should have three elements" do
      expect(resp.roomsArray).to eq(
        [
          {:room=>
            {:elems=>
              {:id=>
                {:elems=>
                  [{:name=>{:text=>"Conf Room 3/101 (16) AV"}},
                   {:email_address=>{:text=>"cf3101@contoso.com"}},
                   {:routing_type=>{:text=>"SMTP"}},
                   {:mailbox_type=>{:text=>"Mailbox"}}]}}}},
          {:room=>
            {:elems=>
              {:id=>
                {:elems=>
                  [{:name=>{:text=>"Conf Room 3/102 (8) AV"}},
                   {:email_address=>{:text=>"cf3102@contoso.com"}},
                   {:routing_type=>{:text=>"SMTP"}},
                   {:mailbox_type=>{:text=>"Mailbox"}}]}}}},
          {:room=>
            {:elems=>
              {:id=>
                {:elems=>
                  [{:name=>{:text=>"Conf Room 3/103 (14) AV RoundTable"}},
                   {:email_address=>{:text=>"cf3103@contoso.com"}},
                   {:routing_type=>{:text=>"SMTP"}},
                   {:mailbox_type=>{:text=>"Mailbox"}}]}}}}
        ]
      )
    end
  end
end
