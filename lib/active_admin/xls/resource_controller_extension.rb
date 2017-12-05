module ActiveAdmin
  module Xls
    module ResourceControllerExtension
      def self.prepended(base)
        base.send :respond_to, :xls
      end

      def index_with_xls
        super do |format|
          block.call format if block_given?

          format.xls do
            xls = active_admin_config.xls_builder.serialize(collection, view_context)
            send_data(xls,
                      filename: xls_filename,
                      type: Mime::Type.lookup_by_extension(:xls))
          end
        end
      end

      # patch per_page to use the CSV record max for pagination
      # when the format is xls
      def per_page
        if request.format == Mime::Type.lookup_by_extension(:xls)
          return max_per_page if respond_to?(:max_per_page, true)
          active_admin_config.max_per_page
        end

        super
      end

      # Returns a filename for the xls file using the collection_name
      # and current date such as 'my-articles-2011-06-24.xls'.
      def xls_filename
        timestamp = Time.now.strftime('%Y-%m-%d')
        "#{resource_collection_name.to_s.tr('_', '-')}-#{timestamp}.xls"
      end
    end
  end
end
