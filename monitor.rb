# Author: Ben Nagy
# Copyright: Copyright (c) Ben Nagy, 2006-2010.
# License: The MIT License
# (See README.TXT or http://www.opensource.org/licenses/mit-license.php for details.)

require File.dirname(__FILE__) + '/drb_debug_client'
require File.dirname(__FILE__) + '/cpumon'
require 'rubygems'
require 'fileutils'
require 'win32api'
require 'trollop'

OPTS=Trollop::options do
    opt :port, "Port to listen on, default 8888", :type=>:integer, :default=>8888
    opt :debug, "Debug mode", :type=>:boolean
end

class Monitor

    COMPONENT="Monitor"
    VERSION="1.6.0"
    MONITOR_DEFAULTS={
        'timeout'=>25,
        'ignore_exceptions'=>[],
        'kill_dialogs'=>true
    }
    CDB_PATH='"C:\\Program Files\\Debugging Tools for Windows (x86)\\cdb.exe" '
    CPUMON_TICKS=6
    CPUMON_THRESH=0.01
    MONITOR_GRANULARITY=0.5

    # Constants for the dialog killer thread
    BMCLICK=0x00F5
    WM_DESTROY=0x0010
    WM_COMMAND=0x111
    IDOK=1
    IDCANCEL=2
    IDNO=7
    IDCLOSE=8
    GW_ENABLEDPOPUP=0x0006
    # Win32 API definitions for the dialog killer
    FindWindow=Win32API.new("user32.dll", "FindWindow", 'PP','N')
    GetWindow=Win32API.new("user32.dll", "GetWindow", 'LI','I')
    PostMessage=Win32API.new("user32.dll", "PostMessage", 'LILL','I')

    def initialize
        warn "#{COMPONENT}:#{VERSION}: Spawning debug server on #{OPTS[:port]+1}..." if OPTS[:debug]
        system("start cmd /k ruby drb_debug_server.rb -p #{OPTS[:port]+1} #{OPTS[:debug]? ' -d' : ''}")
        start_sweeper_thread
        @debug_client=DebugClient.new('127.0.0.1', OPTS[:port]+1)
        @tick_count=0
    rescue
        warn "#{COMPONENT}:#{VERSION}: #{__method__} #{$!} " if OPTS[:debug]
        raise $!
    end

    def start( app_pid, app_wid, arg_hsh={} )
        warn "#{COMPONENT}:#{VERSION}: Starting to monitor pid #{app_pid}" if OPTS[:debug]
        start_debugger( app_pid )
        raise RuntimeError, "#{COMPONENT}:#{VERSION}: Debugee PID mismatch" unless @debug_client.target_pid==app_pid
        @monitor_args=MONITOR_DEFAULTS.merge( arg_hsh )
        start_dk_thread( app_wid ) if @monitor_args['kill_dialogs']
        start_monitor_thread( app_pid )
    rescue
        warn "#{COMPONENT}:#{VERSION}: #{__method__} #{$!} " if OPTS[:debug]
        reset
        raise $!
    end

    def start_debugger( pid )
        @debugger_uri=@debug_client.start_debugger('pid'=>pid, 'options'=>"-xi ld", 'path'=>CDB_PATH )
        @debugger=DRbObject.new nil, @debugger_uri
        @debugger.puts <<-eos
 !load winext\\msec.dll
 .sympath c:\\localsymbols
 sxe -c ".echo frobozz;r;~;kv;u @eip;!exploitable -m;.echo xyzzy" -c2 ".echo frobozz;r;~;kv;u @eip;!exploitable -m;.echo xyzzy" av
 sxe -c ".echo frobozz;r;~;kv;u @eip;!exploitable -m;.echo xyzzy" -c2 ".echo frobozz;r;~;kv;u @eip;!exploitable -m;.echo xyzzy" sbo
 sxe -c ".echo frobozz;r;~;kv;u @eip;!exploitable -m;.echo xyzzy" -c2 ".echo frobozz;r;~;kv;u @eip;!exploitable -m;.echo xyzzy" ii
 sxe -c ".echo frobozz;r;~;kv;u @eip;!exploitable -m;.echo xyzzy" -c2 ".echo frobozz;r;~;kv;u @eip;!exploitable -m;.echo xyzzy" gp
 sxi e0000001
 sxi e0000002
 eos
        @debugger.sync_dq
        @debugger.puts "g"
        warn "#{COMPONENT}:#{VERSION}: Attached debugger to pid #{pid}" if OPTS[:debug]
    rescue
        warn "#{COMPONENT}:#{VERSION}: #{__method__} #{$!} " if OPTS[:debug]
        raise $!
    end

    def start_sweeper_thread
        @sweeper.kill if @sweeper
        @patterns||=['R:/Temp/**/*.*', 'R:/Temporary Internet Files/**/*.*', 'R:/fuzzclient/~$*.doc', 'C:/msbh0day/~$*.doc']
        @sweeper=Thread.new do
            loop do
                @patterns.each {|pattern|
                    Dir.glob(pattern, File::FNM_DOTMATCH).each {|fn|
                        next if File.directory?(fn)
                        begin
                            FileUtils.rm_f(fn)
                        rescue
                            next # probably still open
                        end
                    }
                }
                sleep 5
            end
        end
    end

    def start_dk_thread( app_wid )
        warn "#{COMPONENT}:#{VERSION}: Starting DK thread against wid #{app_wid}" if OPTS[:debug]
        @dk_thread.kill if @dk_thread
        @dk_thread=Thread.new do
            loop do
                begin
                    # Get any descendant windows which are enabled - alerts, dialog boxes etc
                    child_hwnd=GetWindow.call(app_wid, GW_ENABLEDPOPUP)
                    unless child_hwnd==0
                        PostMessage.call(child_hwnd,WM_COMMAND,IDCANCEL,0)
                        PostMessage.call(child_hwnd,WM_COMMAND,IDNO,0)
                        PostMessage.call(child_hwnd,WM_COMMAND,IDCLOSE,0)
                        PostMessage.call(child_hwnd,WM_COMMAND,IDOK,0)
                        PostMessage.call(child_hwnd,WM_DESTROY,0,0)
                    end
                    # conn_office.rb changes the caption, so this should only detect toplevel dialog boxes
                    # that pop up during open before the main Word window.
                    toplevel_box=FindWindow.call(0, "Microsoft Office Word")
                    unless toplevel_box==0
                        PostMessage.call(toplevel_box,WM_COMMAND,IDCANCEL,0)
                        PostMessage.call(toplevel_box,WM_COMMAND,IDNO,0)
                        PostMessage.call(toplevel_box,WM_COMMAND,IDCLOSE,0)
                        PostMessage.call(toplevel_box,WM_COMMAND,IDOK,0)
                    end
                    sleep(0.5)
                    if @monitor_thread.alive? 
                        print 'o' 
                    else
                        print 'x'
                    end
                rescue
                    sleep(0.5)
                    warn "#{COMPONENT}:#{VERSION}: Error in DK thread: #{$!}"
                    retry
                end
            end
        end
    end

    def start_monitor_thread( pid )
        raise RuntimeError, "#{COMPONENT}:#{VERSION}: Debugger not initialized yet!" unless @debugger
        @monitor_thread.kill if @monitor_thread
        @monitor_thread=Thread.new do
            @running=true
            @tick_count=0
            @cpumon=ProcessCPUMonitor.new( pid )
            warn "#{COMPONENT}:#{VERSION}: Monitor thread started for #{pid}" if OPTS[:debug]
            loop do
                begin
                    @pid=pid
                    raise RuntimeError, "PID Mismatch" unless @pid==@debug_client.target_pid
                    sleep MONITOR_GRANULARITY
                    @tick_count+=1
                    @cpumon.update_rolling_avg
                    check_for_timeout
                    if @debugger.target_running?
                        check_for_idle
                    else
                        debugger_output=@debugger.sync_qc
                        warn "#{COMPONENT}:#{VERSION}: Target #{@debug_client.target_pid} broken..." if OPTS[:debug]
                        if fatal_exception? debugger_output
                            warn "#{COMPONENT}:#{VERSION}: Fatal exception. Killing debugee." if OPTS[:debug]
                            warn debugger_output[-500..-1] if OPTS[:debug]
                            treat_as_fatal( debugger_output )
                        else
                            warn "#{COMPONENT}:#{VERSION}: Broken, but no fatal exception. Ignoring." if OPTS[:debug]
                            warn debugger_output[-500..-1] if OPTS[:debug]
                            @debugger.puts "g" 
                        end
                    end
                rescue
                    @running=false
                    warn "#{COMPONENT}:#{VERSION}: #{__method__} #{$!} Set running to false " if OPTS[:debug]
                    @debug_client.close_debugger if @debugger
                    @debugger=nil
                    Thread.exit
                end
            end
        end
    rescue
        warn "#{COMPONENT}:#{VERSION}: #{__method__} #{$!} " if OPTS[:debug]
        @running=false
        raise $!
    end

    def running?
        @running
    end

    def hang?
        @hang
    end

    def fatal_exception?( output )
        unless output.scan(/frobozz/).length==output.scan(/xyzzy/).length
            raise RuntimeError, "#{COMPONENT}:#{VERSION}:#{__method__}: unfinished exception output."
        end
        return true if output=~/second chance/i
        return false unless output=~/frobozz/
            # Does the most recent exception match none of the ignore regexps?
            exception=output.split(/frobozz/i).last
        @monitor_args['ignore_exceptions'].none? {|ignore_string| exception=~(Regexp.new(eval(ignore_string)))} 
    rescue
        warn "#{COMPONENT}:#{VERSION}: #{__method__} #{$@.join "\n"} " if OPTS[:debug]
        raise $!
    end

    def check_for_idle
        # Only called from within the monitor thread, so the Thread.exit
        # exits @monitor_thread not the whole app
        #
        # CPUMON_TICKS is the minimum number of events for rolling_avg - it
        # returns nil otherwise
        if (avg=@cpumon.rolling_avg( CPUMON_TICKS )) && avg < CPUMON_THRESH
            warn "#{COMPONENT}:#{VERSION}: CPU monitor says 'no'. (average #{avg} at #{CPUMON_TICKS} measures)" if OPTS[:debug]
            debugger_output=@debugger.sync_dq
            if fatal_exception? debugger_output
                warn "#{COMPONENT}:#{VERSION}: Fatal exception after idle" if OPTS[:debug]
                treat_as_fatal( debugger_output )
            else
                warn "#{COMPONENT}:#{VERSION}: No exception after idle" if OPTS[:debug]
                @debug_client.close_debugger if @debugger
                @debugger=nil
                Thread.exit
            end
        end
    rescue
        warn "#{COMPONENT}:#{VERSION}: #{__method__} #{$!} " if OPTS[:debug]
        raise $!
    end

    def check_for_timeout
        # Only called from within the monitor thread, so the Thread.exit
        # exits @monitor_thread not the whole app
        if Time.now - @mark > @monitor_args['timeout']
            warn "CPU: #{@cpumon.rolling_avg( 6 )}"
            warn "#{COMPONENT}:#{VERSION}: Hard Timeout (#{Time.now - @mark}) Exceeded." if OPTS[:debug]
            @hang=true
            debugger_output=@debugger.sync_dq
            if (fatal_exception?( debugger_output ) rescue true)
                warn "#{COMPONENT}:#{VERSION}: Fatal exception after timeout" if OPTS[:debug]
                treat_as_fatal( debugger_output )
            else
                warn "#{COMPONENT}:#{VERSION}: No exception after timeout" if OPTS[:debug]
                @debug_client.close_debugger if @debugger
                @debugger=nil
                Thread.exit
            end
        end
    rescue
        warn "#{COMPONENT}:#{VERSION}: #{__method__} #{$!} " if OPTS[:debug]
        raise $!
    end

    def treat_as_fatal( debugger_output )
        # Only called from within the monitor thread, so the Thread.exit
        # exits @monitor_thread not the whole app
        get_minidump if @monitor_args['minidump']
        @exception_data=debugger_output
        @debug_client.close_debugger if @debugger
        @debugger=nil
        Thread.exit
    rescue
        warn "#{COMPONENT}:#{VERSION}: #{__method__} #{$!} " if OPTS[:debug]
        raise $!
    end

    def exception_data
        @exception_data
    end

    def minidump
        @minidump
    end

    def last_tick
        now=@tick_count
        until @tick_count > now
            unless (@monitor_thread.alive? and running?)
                sleep MONITOR_GRANULARITY and return
            end
        end
    end

    def get_minidump
        warn "#{COMPONENT}:#{VERSION}: Collecting minidump..." if OPTS[:debug]
        #do something
        @debugger.puts ".dump /mFhutwd r:\\fuzzclient\\mini.dmp"
        @debugger.sync
        unless File.exists? "R:/fuzzclient/mini.dmp"
            raise RuntimeError, "#{COMPONENT}-#{VERSION}:#{__method__}: Tried to dump, but couldn't find it!"
        end
        @minidump=File.open( "R:/fuzzclient/mini.dmp", "rb" ) {|io| io.read}
        FileUtils.rm_f( "R:/fuzzclient/mini.dmp" )
    end

    def reset
        # Only called externally - trying to kill a thread from inside itself
        # doesn't seem to work properly
        warn "#{COMPONENT}:#{VERSION}: Reset called, debugger #{@debug_client.debugger_pid rescue 0}, running - #{@running}" if OPTS[:debug]
        @debug_client.close_debugger if @debugger
        Thread.kill( @monitor_thread ) if @monitor_thread
        @debugger=nil
        clear_hang
    rescue
        warn "#{COMPONENT}:#{VERSION}: #{__method__} #{$!} " if OPTS[:debug]
        raise $!
    end

    def new_test( filename )
        warn "#{COMPONENT}:#{VERSION}: Prepping for new test #{filename}" if OPTS[:debug]
        raise "#{COMPONENT}:#{VERSION}: Unable to continue, monitor thread dead!" unless @monitor_thread.alive?
        raise "#{COMPONENT}:#{VERSION}: Unable to continue, no debugger" unless @debugger
        raise "#{COMPONENT}:#{VERSION}: Uncleared exception data!!" if @exception_data
        @mark=Time.now 
        @cpumon.clear_rolling_avg
        @debugger.dq_all
    rescue
        warn "#{COMPONENT}:#{VERSION}: #{__method__} #{$!} " if OPTS[:debug]
        raise $!
    end

    def clear_exception
        @exception_data=nil
    end

    def clear_hang
        @hang=false
    end

    def destroy
        warn "#{COMPONENT}:#{VERSION}: Destroying." if OPTS[:debug]
        @debug_client.destroy_server
    rescue
        warn "#{COMPONENT}:#{VERSION}: #{__method__} #{$!} " if OPTS[:debug]
        raise $!
    ensure
        Process.exit!
    end

end

DRb.start_service( "druby://127.0.0.1:#{OPTS[:port]}", Monitor.new )
DRb.thread.join
