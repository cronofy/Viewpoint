require 'spec_helper'

describe Viewpoint::EWS::Types::CalendarItem do

  describe "#duration" do
    it "returns the duration in seconds" do
      allow_any_instance_of(described_class).to receive(:simplify!)
      calitem = described_class.new(nil, nil)
      allow(calitem).to receive(:duration).and_return("PT0H30M0S")
      expect(calitem.duration_in_seconds).to eql(1800)
    end
  end

  describe "#organizer" do
    it "maps the organizer type passed" do
      calitem = described_class.new(nil, {
        :organizer=>{:elems=>[{:mailbox=>{:elems=>[{:name=>{:text=>"Organizer Name"}}, {:email_address=>{:text=>"organizer@example.com"}}, {:routing_type=>{:text=>"EX"}}]}}]}
      })
      expect(calitem.organizer.name).to eql("Organizer Name")
      expect(calitem.organizer.email).to eql("organizer@example.com")
      expect(calitem.organizer.email_address).to eql("organizer@example.com")
    end
  end

  describe "#my_response_type" do
    it "maps the response type passed" do
      calitem = described_class.new(nil, { my_response_type: { text: "Organizer" }})
      expect(calitem.my_response_type).to eql("Organizer")
    end
  end

end
