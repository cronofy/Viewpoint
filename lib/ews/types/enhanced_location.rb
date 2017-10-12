module Viewpoint::EWS::Types
  class EnhancedLocation
    include Viewpoint::EWS
    include Viewpoint::EWS::Types

    ENHANCED_LOCATION_KEY_PATHS = {
      display_name:   [:display_name],
      annotation:     [:annotation],
      postal_address: [:postal_address],
    }
    ENHANCED_LOCATION_KEY_TYPES = {
      postal_address: :build_postal_address,
    }
    ENHANCED_LOCATION_KEY_ALIAS = { }

    def initialize(ews, enhanced_location)
      @ews = ews
      @ews_item = enhanced_location
      simplify!
    end

    def build_postal_address(postal_address_ews)
      Types::PostalAddress.new(ews, postal_address_ews)
    end

    private

    def simplify!
      @ews_item = @ews_item.inject({}){|m,o|
        m[o.keys.first] = o.values.first[:text] || o.values.first[:elems];
        m
      }
    end

    def key_paths
      @key_paths ||= super.merge(ENHANCED_LOCATION_KEY_PATHS)
    end

    def key_types
      @key_types ||= super.merge(ENHANCED_LOCATION_KEY_TYPES)
    end

    def key_alias
      @key_alias ||= super.merge(ENHANCED_LOCATION_KEY_ALIAS)
    end

  end # EnhancedLocation
end # Viewpoint::EWS::Types
