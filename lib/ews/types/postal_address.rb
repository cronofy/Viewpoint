module Viewpoint::EWS::Types
  class PostalAddress
    include Viewpoint::EWS
    include Viewpoint::EWS::Types

    POSTAL_ADDRESS_KEY_PATHS = {
      street:            [:street],
      city:              [:city],
      state:             [:state],
      country:           [:country],
      postal_code:       [:postal_code],
      type:              [:type],
      latitude:          [:latitude],
      longitude:         [:longitude],
      accuracy:          [:accuracy],
      altitude:          [:altitude],
      altitude_accuracy: [:altitude_accuracy],
      formatted_address: [:formatted_address],
      location_uri:      [:location_uri],
      location_source:   [:location_source],
    }
    POSTAL_ADDRESS_KEY_TYPES = { }
    POSTAL_ADDRESS_KEY_ALIAS = { }

    def initialize(ews, postal_address)
      @ews = ews
      @ews_item = postal_address
      simplify!
    end

    private

    def simplify!
      @ews_item = @ews_item.inject({}){|m,o|
        m[o.keys.first] = o.values.first[:text] || o.values.first[:elems];
        m
      }
    end

    def key_paths
      @key_paths ||= super.merge(POSTAL_ADDRESS_KEY_PATHS)
    end

    def key_types
      @key_types ||= super.merge(POSTAL_ADDRESS_KEY_TYPES)
    end

    def key_alias
      @key_alias ||= super.merge(POSTAL_ADDRESS_KEY_ALIAS)
    end

  end # PostalAddress
end # Viewpoint::EWS::Types
