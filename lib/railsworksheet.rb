require 'yajb/jbridge'
require 'yajb/jlambda'
require 'tempfile'

PATH_SEPARATOR =  RUBY_PLATFORM.include?("win32") ? ';' : ':' 
    
#It would be more readable to keep the jar file name with full version details
#So just look for the pattern in the jar file name ignoring the version details

POI_JAR = Dir[File.join(File.dirname(__FILE__), 'poi*jar')][0] 

JBRIDGE_OPTIONS = {
   :jvm_stdout => :never,
   :classpath =>  POI_JAR + PATH_SEPARATOR + File.dirname(__FILE__) 
}

include JavaBridge

class WorksheetRender
    	include ApplicationHelper
      include ActionView::Helpers::NumberHelper

      def self.compilable?
           false
         end

      def compilable?
        false
      end

    	def initialize(action_view)
      		@action_view = action_view
    	end
    
	    def render(template, local_assigns = {})
	    	
	    	
	    	#get the instance variables setup	    	
      	@action_view.controller.instance_variables.each do |v|
        		instance_variable_set(v, @action_view.controller.instance_variable_get(v))
		    end
			  @rails_worksheet_name = "Default.xls" if @rails_worksheet_name.nil?
				
    		@action_view.controller.headers["Content-Type"] ||= 'application/xls'
			  @action_view.controller.headers["Content-Disposition"] ||= 'attachment; filename="' + @rails_worksheet_name + '"'      
        
        jimport "java.io.*"
        jimport "org.apache.poi.hssf.usermodel.*"
        workbook = jnew :HSSFWorkbook

        CellBatch.instance.clear
        RowGroupBatch.instance.clear
                
        eval template.source, nil, "#{@action_view.base_path}/#{@action_view.first_render}.#{@action_view.finder.pick_template_extension(@action_view.first_render)}" 


        # Save two lines each in the template where the user is building only
        # worksheet. Will do no harm if the stuff is already written out.
            
        CellBatch.write_to(workbook.getSheetAt(0)) if CellBatch.instance.write_pending?
        RowGroupBatch.group_rows(workbook.getSheetAt(0)) if RowGroupBatch.instance.group_pending?

        begin
          temp = Tempfile.new('railsworksheet-', File.join(RAILS_ROOT, 'tmp') )
          out = jnew :FileOutputStream, temp.path
          workbook.write(out)
          out.close
          File.open(temp.path, 'rb') { |file| file.read }
        ensure
          File.delete(temp.path)
        end
    end
        
end
