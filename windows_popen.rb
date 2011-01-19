# Author: Ben Nagy
# Copyright: Copyright (c) Ben Nagy, 2006-2010.
# License: The MIT License
# (See README.TXT or http://www.opensource.org/licenses/mit-license.php for details.)

# Mostly stolen from Simon Kroger, comp.lang.ruby post Sep 12, 2005.
# I fixed the handle leaks and also added the CREATE_NEW_PROCESS_GROUP stuff
# which allows me to send a break to the debugger.

require 'win32/pipe'
require 'win32/process'
module WindowsPipe
    require 'Win32API' 
    include Windows::Pipe
    include Windows::Handle
    include Windows::Process
    include Windows::File
    include Windows::Error
    extend self
    $stdout.sync = true 
    NORMAL_PRIORITY_CLASS = 0x00000020 
    CREATE_NEW_PROCESS_GROUP=0x00000200
    STARTUP_INFO_SIZE = 68 
    PROCESS_INFO_SIZE = 16 
    SECURITY_ATTRIBUTES_SIZE = 12 
    ERROR_SUCCESS = 0x00 
    FORMAT_MESSAGE_FROM_SYSTEM = 0x1000 
    FORMAT_MESSAGE_ARGUMENT_ARRAY = 0x2000 
    HANDLE_FLAG_INHERIT = 1 
    HANDLE_FLAG_PROTECT_FROM_CLOSE =2 
    STARTF_USESHOWWINDOW = 0x00000001 
    STARTF_USESTDHANDLES = 0x00000100 

    def raise_last_win_32_error 
        errorCode=GetLastError.call
        if errorCode != ERROR_SUCCESS 
            msg = ' ' * 255 
            msgLength = FormatMessage.call(0x3000, 0, errorCode, 0, msg, 255, '') 
            msg.gsub!(/\000/, '') 
            msg.strip! 
            raise "WindowsPipe: Win32 Exception: #{msg}" 
        else 
            raise 'GetLastError returned ERROR_SUCCESS' 
        end 
    end 

    def create_pipe # returns read and write handle 
        read_handle, write_handle = [0].pack('I'), [0].pack('I') 
        sec_attrs = [SECURITY_ATTRIBUTES_SIZE, 0, 1].pack('III') 
        raise_last_win_32_error if CreatePipe.call(read_handle, write_handle, sec_attrs, 0).zero? 
        [read_handle.unpack('I')[0], write_handle.unpack('I')[0]] 
    end 

    def set_handle_information(handle, flags, value) 
        raise_last_win_32_error if SetHandleInformation.call(handle, flags, value).zero? 
    end 

    def close_handle(handle) 
        raise_last_win_32_error if CloseHandle.call(handle).zero? 
    end 

    def create_process(command, stdin, stdout, stderror) 
        startupInfo = [STARTUP_INFO_SIZE, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 
            STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW, 0, 
            0, 0, stdin, stdout, stderror].pack('IIIIIIIIIIIISSIIII') 
        processInfo = [0, 0, 0, 0].pack('IIII') 
        command << "\x00"
        # The CREATE_NEW_PROCESS_GROUP flag allows me to send console events
        # to the process group with kernel32!GenerateConsoleCtrlEvent
        # which is important if you want to send CTRL-BREAK to a debugger.
        raise_last_win_32_error if CreateProcess.call(0, command, 0, 0, 1, 
                                                      CREATE_NEW_PROCESS_GROUP, 0, 
                                                      0, startupInfo, processInfo).zero? 
        hProcess, hThread, dwProcessId, dwThreadId = processInfo.unpack('LLLL') 
        # Close the process and thread handles, so they don't hang around.
        # Users can always open new ones if they need them.
        close_handle(hProcess) 
        close_handle(hThread) 
        [dwProcessId, dwThreadId] 
    end 

    def write_file(hFile, buffer) 
        written = [0].pack('I') 
        raise_last_win_32_error if WriteFile.call(hFile, buffer, buffer.size, written, 0).zero? 
        written.unpack('I')[0] 
    end 

    def read_file( hFile ) 
        number = [0].pack('I') 
        buffer = ' ' * 255 
        return '' if ReadFile.call(hFile, buffer, 255, number, 0).zero? 
        buffer[0...number.unpack('I')[0]] 
    end 

    def peek_named_pipe(hFile) 
        available = [0].pack('I') 
        return -1 if PeekNamedPipe.call(hFile, 0, 0, 0, available, 0).zero? 
        available.unpack('I')[0] 
    end 

    class Win32popenIO 
        attr_reader :pid

        def initialize (hRead, hWrite, pid) 
            extend WindowsPipe
            @hRead = hRead 
            @hWrite = hWrite 
            @pid=pid
        end 

        def write data 
            write_file(@hWrite, data.to_s) 
        end 

        def read 
            sleep(0.01) while peek_named_pipe(@hRead).zero? 
            read_file(@hRead) 
        end 

        def read_all 
            all = '' 
            while !(buffer = read).empty? 
                all << buffer 
            end 
            all 
        end 

        def close
            close_handle(@hRead) rescue nil
            close_handle(@hWrite) rescue nil
        end

    end 
    module_function
    def popen(command, wrap_in_shell=false) 
        # create 3 pipes 
        child_in_r, child_in_w = create_pipe 
        child_out_r, child_out_w = create_pipe 
        child_error_r, child_error_w = create_pipe 
        # Ensure the write handle to the pipe for STDIN is not inherited. 
        set_handle_information(child_in_w, HANDLE_FLAG_INHERIT, 0) 
        set_handle_information(child_out_r, HANDLE_FLAG_INHERIT, 0) 
        set_handle_information(child_error_r, HANDLE_FLAG_INHERIT, 0) 
        if wrap_in_shell
            processId, threadId = create_process(ENV['ComSpec'] + ' /C ' + 
                                                 command, child_in_r, child_out_w, child_error_w) 
        else
            processId, threadId = create_process(command, child_in_r, child_out_w, child_error_w) 
        end
        # we have to close the handles, so the pipes terminate with the process 
        # the original code didn't close ALL the handles, so it leaked pipes :(
        close_handle(child_in_r) 
        close_handle(child_out_w) 
        close_handle(child_error_w)
        close_handle(child_error_r) 
        Win32popenIO.new(child_out_r, child_in_w, processId) 
    end 
end # module WindowsPipe
if $0 == __FILE__ 
    io = WindowsPipe.popen('ver',true) 
    puts io.read_all 
    io = WindowsPipe.popen('cmd') 
    io.write("ping localhost\nexit\n") 
    while !(buffer = io.read).empty? 
        print buffer.gsub("\r", '') 
    end 
end 

