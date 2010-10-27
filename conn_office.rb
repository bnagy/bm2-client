# Author: Ben Nagy
# Copyright: Copyright (c) Ben Nagy, 2006-2010.
# License: The MIT License
# (See README.TXT or http://www.opensource.org/licenses/mit-license.php for details.)

require 'win32ole'
require 'win32api'
require 'win32/process'

#Send data to an Office application via file, used for file fuzzing.
#
#Currently the calling code is expected to manage the files, so the deliver method takes a filename as
#its parameter
module CONN_OFFICE

    include Windows::Error
    include Windows::Window
    include Windows::Process
    include Windows::Handle

    #These methods will override the stubs present in the Connector
    #class, and implement the protocol specific functionality for 
    #these generic functions.
    #
    #Arguments required to set up the connection are stored in the
    #Connector instance variable @module_args.
    #
    #Errors should be handled at the Module level (ie here), since Connector
    #just assumes everything is going to plan.
    #
    def raise_win32_error 
        unless (err_code=GetLastError.call)==ERROR_SUCCESS 
            msg = ' ' * 255 
            FormatMessage.call(0x3000, 0, err_code, 0, msg, 255, '') 
            msg.gsub!(/\000/, '').strip! 
            raise "CONN_OFFICE: Win32 Exception: #{msg}" 
        else 
            raise 'GetLastError returned ERROR_SUCCESS' 
        end 
    end 

    #Open the application via OLE
    def pid_from_app(win32ole_app)
        # This approach is straight from MS docs, but it's a horrible hack. Set the window title
        # so we can tell it apart from any other Word instances, find the hWND, then use that
        # to find the PID. Will collide if another window has the same random number.
        win32ole_app.caption=(cookie=rand(2**32).to_s)
        raise_win32_error if ( wid=FindWindow.call( 0, cookie ) ).zero? 
        pid=[0].pack('L') #will be filled in, because it's passed as a pointer
        raise_win32_error if ( GetWindowThreadProcessId.call( wid, pid ) ).zero? 
        [ pid.unpack('L').first , wid ]
    end
    private :pid_from_app
    attr_reader :pid,:wid

    #Open the application via OLE	
    def establish_connection
        @appname = @module_args[0]
        begin
            @app=WIN32OLE.new(@appname+'.Application')
            @app.visible=false
            @pid,@wid=pid_from_app( @app )
            @hprocess=OpenProcess.call( PROCESS_TERMINATE, 0, @pid )
            @app.DisplayAlerts=0
        rescue
            close
            raise RuntimeError, "CONN_OFFICE: establish: couldn't open application. (#{$!})"
        end
    end

    # Don't know what this could be good for...
    def blocking_read
        ''
    end

    # Take a filename and open it in the application
    def blocking_write( filename )
        raise RuntimeError, "CONN_OFFICE: blocking_write: Not connected!" unless is_connected?
        begin
            # this call blocks, so if it opens a dialog box immediately we lose control of the app. 
            # This is the biggest issue, and so far can only be solved with a separate monitor app
            @app.Documents.Open({"FileName"=>filename,"AddToRecentFiles"=>false,"Visible"=>false})
        rescue
            raise RuntimeError, "CONN_OFFICE: blocking_write: Couldn't write to application! (#{$!})"
        end
    end

    def visible=( bool )
        @app.visible=bool
    end

    #Return a boolen.
    def is_connected?
        begin
            @app.visible # any OLE call will fail if the app has died
            return true  
        rescue
            return false
        end		
    end

    def dialog_boxes?
        # 0x06 == GW_ENABLEDPOPUP, which is for subwindows that have grabbed focus.
        (not GetWindow.call( @wid, 0x06 ).zero?) rescue false
    end

    def close_documents
        return true unless @app
        until @app.Documents.count==0
            @app.ActiveDocument.close rescue break
        end
        # Sometimes ruby needs a helping hand.
        GC.start
    end

    def destroy_connection
        begin
            sleep(0.1) while dialog_boxes?
            begin
                @app.Quit if is_connected?
            rescue
                TerminateProcess.call(@hprocess,1)
            end
        ensure
            @app.ole_free rescue nil
            CloseHandle.call( @hprocess )
            @app=nil
        end
    end

end
