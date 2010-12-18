# Author: Ben Nagy
# Copyright: Copyright (c) Ben Nagy, 2006-2010.
# License: The MIT License
# (See README.TXT or http://www.opensource.org/licenses/mit-license.php for details.)

require File.dirname(__FILE__) + '/windows_popen'
require 'win32api'
require 'win32/process'

#Establish a connection to the Windows CDB debugger. CDB has all the features of WinDbg, but it uses
#a simple command line interface.
#
#Parameters: For now, the full command line EXCLUDING the path to cdb itself as a string. 
#Sugar may follow later. Remember that program names that include spaces need to be 
#enclosed in quotes \"c:\\Program Files [...] \" etc.
module CONN_CDB

    CDB_PATH="\"C:\\WinDDK\\Debuggers\\cdb.exe\" "
    COMPONENT="CONN_CDB"
    VERSION="3.6.0"
    include Windows::Thread
    include Windows::Handle
    include Windows::Error
    include Windows::ToolHelper
    include Windows::Process
    GenerateConsoleCtrlEvent=Win32API.new("kernel32", "GenerateConsoleCtrlEvent", ['I','I'], 'I')
    # Used for reliability testing
    INJECT_FAULTS=false
    FAULT_CHANCE=5

    def raise_win32_error( str="" ) 
        unless (error_code=GetLastError.call) == ERROR_SUCCESS 
            msg = ' ' * 255 
            FormatMessage.call(0x3000, 0, error_code, 0, msg, 255, '') 
            raise "#{COMPONENT}:#{VERSION}: #{str} Win32 Exception: #{msg.gsub!(/\000/, '').strip!}" 
        else 
            raise 'GetLastError returned ERROR_SUCCESS' 
        end 
    end

    def establish_connection
        arg_hash=@module_args[0]
        raise ArgumentError, "CONN_CDB: No Pid to attach to!" unless arg_hash['pid']
        @target_pid=arg_hash['pid']
        begin
            command=(arg_hash['path'] || CDB_PATH)+"-p #{arg_hash['pid']} "+"#{arg_hash['options']}"
            @cdb_app=WindowsPipe.popen( (arg_hash['path'] || CDB_PATH)+"-p #{arg_hash['pid']} "+"#{arg_hash['options']}" )
        rescue
            $sterr.puts $!
            $sterr.puts $@
            @cdb_app.close
            raise $!
        end
    end

    # Return the pid of the debugger
    def debugger_pid
        if @cdb_app
            @cdb_app.pid
        else
            -1
        end
    end

    def target_pid
        @target_pid||=false
        @target_pid
    end

    #Blocking read from the socket.
    def blocking_read
        @cdb_app.read
    end

    #Blocking write to the socket.
    def blocking_write( data )
        @cdb_app.write data
    end

    #Return a boolen.
    def is_connected?
        # The process is alive - is that the same as connected?
        begin
            return false if (hProcess=OpenProcess.call( PROCESS_QUERY_INFORMATION, 0, @cdb_pid )).zero?
            return true
        ensure
            CloseHandle.call hProcess 
        end
    end

    #Cleanly destroy the socket. 
    def destroy_connection
        begin
            # Don't use Process.kill( 1,... here because that creates a remote
            # thread, which ends up leaking thread handles when the process is 
            # suspended and then the debugger exits.
            @cdb_app.close if @cdb_app
            # Right now, windows kills CDB when the last handle to it is
            # closed, which also kills the target.
            @cdb_app=nil # for if destroy_connection gets called twice
        rescue
            $stderr.puts $!
            $stderr.puts $@
            raise $!
        end
    end

    # Sugar from here on.

    #Our popen object isn't actually an IO obj, so it only has read and write.
    def puts( str )
        blocking_write "#{str}\n"
    end

    def send_break
        # 1 -> Ctrl-Break event
        GenerateConsoleCtrlEvent.call( 1, @cdb_app.pid )
        sleep(0.1)
    end

    def sync
        raise "FAULT" if INJECT_FAULTS && rand(100)+1 > (100-FAULT_CHANCE)
        # This is awful.
        send_break if target_running?
        puts ".echo #{cookie=rand(2**32)}"
        mark=Time.now
        until qc_all =~ /#{cookie}/
            sleep 1
            raise "#{COMPONENT}:#{VERSION}:#{__method__}: Timed out" if Time.now - mark > 3
        end
    end

    # Here to reduce calls when using over DRb
    def sync_dq
        sync
        dq_all
    end

    # Here to reduce calls when using over DRb
    def sync_qc
        sync
        qc_all
    end

    def target_running?
        raise "FAULT" if INJECT_FAULTS && rand(100)+1 > (100-FAULT_CHANCE)
        # If there's no target pid then we're being called weirdly, but
        # it's definitely not running.
        return false unless target_pid
        begin
            # If we can't get a handle to it, it's not running, no need to do the
            # expensive Toolhelp stuff.
            return false if (hProcess=OpenProcess.call( PROCESS_QUERY_INFORMATION, 0, target_pid )).zero?
        ensure
            CloseHandle.call hProcess 
        end
        # General approach is to check all the target threads. If any
        # are running, then the app is running. If they are all suspended
        # then it's frozen by the debugger. The only good way to check
        # if a thread is suspended is to suspend it, which returns the
        # suspend count - 0 if it was previously running.
        begin
            raise_win32_error("CreateSnap") if (hSnap=CreateToolhelp32Snapshot.call( TH32CS_SNAPTHREAD, 0 ))==INVALID_HANDLE_VALUE
            # I'm going to go ahead and do this the horrible way. This is a
            # blank Threadentry32 structure, with the size (28) as the first
            # 4 bytes (little endian). It will be filled in by the Thread32Next
            # calls
            thr_raw="\x1c" << "\x00"*27
            raise_win32_error("Thread32First") unless Thread32First.call(hSnap, thr_raw)==1
            while Thread32Next.call(hSnap, thr_raw)==1
                # Again, manually 'parsing' the structure in hideous fashion
                owner=thr_raw[12..15].unpack('L').first
                tid=thr_raw[8..11].unpack('L').first
                if owner==target_pid
                    begin
                        raise_win32_error("OpenThread #{tid}") if (hThread=OpenThread.call( THREAD_SUSPEND_RESUME,0,tid )).zero?
                        retry_count=5
                        while (suspend_count=SuspendThread.call( hThread ))==INVALID_HANDLE_VALUE
                            unless (retry_count-=1)<=0
                                sleep(0.1)
                            else
                                raise_win32_error "SuspendThread"
                            end
                        end
                        raise_win32_error("ResumeThread") if (ResumeThread.call( hThread ))==INVALID_HANDLE_VALUE
                    ensure
                        raise_win32_error("CloseHandle") if CloseHandle.call( hThread ).zero?
                    end
                    return true if suspend_count==0
                end
            end
            return false
        ensure
            CloseHandle.call( hSnap )
        end
    end

    def registers
        raise "FAULT" if INJECT_FAULTS && rand(100)+1 > (100-FAULT_CHANCE)
        send_break if target_running?
        puts 'r'
        sync 
        mark=Time.now
        until (regstring=qc_all) =~ /eax.*?efl=.*$/m
            raise "#{COMPONENT}:#{VERSION}:#{__method__}: #{$!}" if Time.now - mark > 5
        end
        Hash[*(regstring.scan(/eax.*?efl=.*$/m).last.scan(/ (\w+?)=([0-9a-f]+)/)).flatten]
    rescue
        $stderr.puts "#{COMPONENT}:#{VERSION}: #{__method__} #{$!} "
        raise $!
    end

end # module CONN_CDB

