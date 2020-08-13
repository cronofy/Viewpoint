module Viewpoint::EWS::Types
  class CalendarItem
    include Viewpoint::EWS
    include Viewpoint::EWS::Types
    include Viewpoint::EWS::Types::Item
    include Viewpoint::StringUtils

    CALENDAR_ITEM_FIELD_URIS = {
      :recurrence  => {:text => 'calendar:Recurrence', :writable => true},
    }

    CALENDAR_ITEM_FIELD_NESTED_UPDATES = %i{
      recurrence
    }

    CALENDAR_ITEM_KEY_PATHS = {
      recurring?:   [:is_recurring, :text],
      meeting?:     [:is_meeting, :text],
      cancelled?:   [:is_cancelled, :text],
      duration:     [:duration, :text],
      time_zone:    [:time_zone, :text],
      reminder_set?: [:is_reminder_set, :text],
      reminder_minutes_before_start: [:reminder_minutes_before_start, :text],
      start:        [:start, :text],
      end:          [:end, :text],
      location:     [:location, :text],
      all_day?:     [:is_all_day_event, :text],
      uid:        [:u_i_d, :text],
      legacy_free_busy_status: [:legacy_free_busy_status, :text],
      my_response_type:   [:my_response_type, :text],
      organizer: [:organizer, :elems, 0, :mailbox, :elems],
      optional_attendees: [:optional_attendees, :elems ],
      required_attendees: [:required_attendees, :elems ],
      recurrence: [:recurrence, :elems ],
      deleted_occurrences: [:deleted_occurrences, :elems ],
      modified_occurrences: [:modified_occurrences, :elems ],
      calendar_item_type: [:calendar_item_type, :text ],
      enhanced_location: [:enhanced_location, :elems ],
      start_time_zone_id: [:start_time_zone_id, :text],
      end_time_zone_id: [:end_time_zone_id, :text],
      join_online_meeting_url: [:join_online_meeting_url, :text ],
   }

    CALENDAR_ITEM_KEY_TYPES = {
      start:        ->(str){DateTime.parse(str)},
      end:          ->(str){DateTime.parse(str)},
      recurring?:   ->(str){str.downcase == 'true'},
      meeting?:     ->(str){str.downcase == 'true'},
      cancelled?:   ->(str){str.downcase == 'true'},
      all_day?:     ->(str){str.downcase == 'true'},
      reminder_set?: ->(str){str.downcase == 'true'},
      reminder_minutes_before_start: ->(str){str.to_i},
      organizer: :build_mailbox_user,
      optional_attendees: :build_attendees_users,
      required_attendees: :build_attendees_users,
      deleted_occurrences: :build_deleted_occurrences,
      modified_occurrences: :build_modified_occurrences,
      enhanced_location: :build_enhanced_location,
    }

    CALENDAR_ITEM_KEY_ALIAS = {}

    # Updates the specified item attributes
    #
    # Uses `SetItemField` if value is present and `DeleteItemField` if value is nil
    # @param updates [Hash] with (:attribute => value)
    # @param options [Hash]
    # @option options :conflict_resolution [String] one of 'NeverOverwrite', 'AutoResolve' (default) or 'AlwaysOverwrite'
    # @option options :send_meeting_invitations_or_cancellations [String] one of 'SendToNone' (default), 'SendOnlyToAll',
    #   'SendOnlyToChanged', 'SendToAllAndSaveCopy' or 'SendToChangedAndSaveCopy'
    # @return [CalendarItem, false]
    # @example Update Subject and Body
    #   item = #...
    #   item.update_item!(subject: 'New subject', body: 'New Body')
    # @see http://msdn.microsoft.com/en-us/library/exchange/aa580254.aspx
    # @todo AppendToItemField updates not implemented
    def update_item!(updates, options = {})
      item_updates = []
      updates.each do |attribute, value|
        if CALENDAR_ITEM_FIELD_URIS.include?(attribute)
          item_field = CALENDAR_ITEM_FIELD_URIS[attribute][:text]
        elsif FIELD_URIS.include?(attribute)
          item_field = FIELD_URIS[attribute][:text]
        end

        field = {field_uRI: {field_uRI: item_field}}

        if value.nil? && item_field
          # Build DeleteItemField Change
          item_updates << {delete_item_field: field}
        elsif value.is_a?(Array) && value.empty?
          item_updates << {delete_item_field: field}
        elsif attribute == :required_attendees
          # Updating property
          elements = value.map do |attendee|
            mailbox = attendee[:attendee][:mailbox]

            elements = []
            elements << { "Name" => { text: mailbox[:name] } } if mailbox[:name]
            elements << { "EmailAddress" => { text: mailbox[:email_address] } } if mailbox[:email_address]

            mailbox = [
              {
                "Mailbox" => {
                  sub_elements: elements
                }
              }
            ]
            { "Attendee" => { sub_elements: mailbox } }
          end

          item_attributes = { "RequiredAttendees" => { sub_elements: elements } }
          item_updates << {set_item_field: field.merge(calendar_item: {sub_elements: item_attributes})}
        elsif attribute == :enhanced_location
          if value[:value] == :delete
            # Deleting property
            item_updates << { delete_item_field: { field_uRI: { field_uRI: "calendar:EnhancedLocation"} } }
          else
            # Updating property
            elements = []
            elements << { "DisplayName" => { text: value[:display_name] } } if value[:display_name]
            elements << { "Annotation" => { text: value[:annotation] } } if value[:annotation]

            if address = value[:postal_address]
              address_elements = []

              address_elements << { "Street" => { text: address[:street] } } if address[:street]
              address_elements << { "City" => { text: address[:city] } } if address[:city]
              address_elements << { "State" => { text: address[:state] } } if address[:state]
              address_elements << { "Country" => { text: address[:country] } } if address[:country]
              address_elements << { "PostalCode" => { text: address[:postal_code] } } if address[:postal_code]
              address_elements << { "Type" => { text: address[:type] } } if address[:type]
              address_elements << { "Latitude" => { text: address[:latitude] } } if address[:latitude]
              address_elements << { "Longitude" => { text: address[:longitude] } } if address[:longitude]
              address_elements << { "Accuracy" => { text: address[:accuracy] } } if address[:accuracy]
              address_elements << { "Altitude" => { text: address[:altitude] } } if address[:altitude]
              address_elements << { "AltitudeAccuracy" => { text: address[:altitude_accuracy] } } if address[:altitude_accuracy]
              address_elements << { "FormattedAddress" => { text: address[:formatted_address] } } if address[:formatted_address]
              address_elements << { "LocationUri" => { text: address[:location_uri] } } if address[:location_uri]
              address_elements << { "LocationSource" => { text: address[:location_source] } } if address[:location_source]

              elements << { "PostalAddress" => { sub_elements: address_elements } }
            end

            item_attributes = {
              "EnhancedLocation" => {
                sub_elements: elements
              }
            }

            item_updates << {set_item_field: field.merge(calendar_item: {sub_elements: item_attributes})}
          end
        elsif CALENDAR_ITEM_FIELD_NESTED_UPDATES.include?(attribute)
          # Build SetItemField Change
          item = Viewpoint::EWS::Template::CalendarItem.new(attribute => value)

          # Remap attributes because ews_builder #dispatch_field_item! uses #build_xml!
          item_attributes = item.to_ews_item.map do |name, value|
            case value
            when String
              { name => { text: value } }
            when Hash
              { name => { sub_elements: convert_update_to_sub_elements(Viewpoint::EWS::SOAP::EwsBuilder.camel_case_attributes(value)) } }
            else
              { name => value }
            end
          end

          item_updates << {set_item_field: field.merge(calendar_item: {sub_elements: item_attributes})}

        elsif attribute == :attachments
          log.debug { "Attachment update - item_field=#{item_field.inspect} attribute=#{attribute} value=#{value}" }

          file_attachments = value.map do |att|
            log.debug { "Parsing attachment - att=#{att[:id]}" }

            {
              'FileAttachment' => {
                sub_elements: [
                  { 'Name' => { text: att[:name] } },
                  { 'Content' => { text: att[:content] } },
                  { 'ContentId' => { text: att[:id] } },
                ]
              }
            }
          end

          update_attachments = {
            set_item_field: field.merge({
              calendar_item: {
                sub_elements: {
                  'Attachments' => { sub_elements: file_attachments }
                }
              }
            })
          }

          item_updates << update_attachments

        elsif item_field
          log.debug { "ItemField update - item_field=#{item_field.inspect} attribute=#{attribute} value=#{value}"}

          # Build SetItemField Change
          item = Viewpoint::EWS::Template::CalendarItem.new(attribute => value)

          # Remap attributes because ews_builder #dispatch_field_item! uses #build_xml!
          item_attributes = item.to_ews_item.map do |name, value|
            case value
            when String
              {name => {text: value}}
            when Hash
              {name => Viewpoint::EWS::SOAP::EwsBuilder.camel_case_attributes(value)}
            else
              {name => value}
            end
          end

          item_updates << {set_item_field: field.merge(calendar_item: {sub_elements: item_attributes})}
        elsif attribute == :extended_property
          values = [value].flatten

          values.each do |value|
            if value[:value] == :delete
              # Deleting property
              item_updates << { delete_item_field: { extended_field_uri: value[:extended_field_uri] } }
            else
              # Updating property
              item_attributes = {
                "ExtendedProperty" => {
                  sub_elements: [
                    {
                      "ExtendedFieldURI" => {
                        "DistinguishedPropertySetId" => value[:extended_field_uri][:distinguished_property_set_id],
                        "PropertyName" => value[:extended_field_uri][:property_name],
                        "PropertyType" => value[:extended_field_uri][:property_type],
                      },
                    },
                    {
                      "Value" => {
                        text: value[:value],
                      },
                    },
                  ]
                }
              }

              item_updates << { set_item_field: { extended_field_uri: value[:extended_field_uri], calendar_item: { sub_elements: item_attributes } } }
            end
          end
        else
          # Ignore unknown attribute
          log.debug { "Passed unknown attribute - #{attribute}" }
        end
      end

      if item_updates.any?
        data = {}
        data[:conflict_resolution] = options[:conflict_resolution] || 'AutoResolve'
        data[:send_meeting_invitations_or_cancellations] = options[:send_meeting_invitations_or_cancellations] || 'SendToNone'
        data[:item_changes] = [{item_id: self.item_id, updates: item_updates}]
        rm = ews.update_item(data).response_messages.first
        if rm && rm.success?
          self.get_all_properties!
          self
        else
          if rm
            raise EwsCreateItemError, "Could not update calendar item. #{rm.code}: #{rm.message_text}"
          else
            raise EwsCreateItemError, "Could not update calendar item."
          end
        end
      end
    end

    def duration_in_seconds
      iso8601_duration_to_seconds(duration)
    end

    def build_enhanced_location(location_ews)
      Types::EnhancedLocation.new(ews, location_ews)
    end

    private

    def key_paths
      super.merge(CALENDAR_ITEM_KEY_PATHS)
    end

    def key_types
      super.merge(CALENDAR_ITEM_KEY_TYPES)
    end

    def key_alias
      super.merge(CALENDAR_ITEM_KEY_ALIAS)
    end

    def convert_update_to_sub_elements(hash)
      hash.map do |key, value|
        case value
        when Hash
          if value[:text]
            { key => value }
          elsif value[:sub_elements]
            { key => value }
          else
            { key => { sub_elements: convert_update_to_sub_elements(value) } }
          end
        else
          { key => { text: value.to_s } }
        end
      end
    end
  end
end
