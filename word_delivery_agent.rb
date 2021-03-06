# Author: Ben Nagy
# Copyright: Copyright (c) Ben Nagy, 2006-2010.
# License: The MIT License
# (See README.TXT or http://www.opensource.org/licenses/mit-license.php for details.)

require File.dirname(__FILE__) + '/../bm2-core/connector'
require File.dirname(__FILE__) + '/conn_office'
require File.dirname(__FILE__) + '/drb_debug_client'
require File.dirname(__FILE__) + '/process_killer'

class WordDeliveryAgent

    COMPONENT="WordDeliveryAgent"
    VERSION="1.6.1"
    DELIVERY_DEFAULTS={
        'clean'=>false, 
        'filechain'=>false,
        'maxchain'=>15,
        'ignore_exceptions'=>[]
    }

    AGENT_DEFAULTS={
        'debug'=>false,
        'visible'=>true
    }
    RETRIES=50

    def debug_info( str )
        warn "#{COMPONENT}-#{VERSION}: #{str}" if @agent_options['debug']
    end

    def initialize( arg_hash={} )
        ProcessKiller::nicely_kill "WINWORD.EXE"
        # For some reason explorer.exe can't be killed with Process.kill(9)
        # it just restarts again :/
        begin
            explorer_pid=ProcessKiller::pids("explorer.exe").first
            hprocess=Windows::Process::OpenProcess.call(Windows::Process::PROCESS_TERMINATE,0,explorer_pid)
            Windows::Process::TerminateProcess.call(hprocess,1)
        rescue
            #can't be bothered
        end
        @agent_options=AGENT_DEFAULTS.merge( arg_hash )
        # Start with high priority for better chance of killing processes pegging the CPU
        debug_info "Starting monitor server..."
        system("start cmd /k ruby monitor.rb #{(@agent_options['debug']? '-d' : ' ')}")
        @monitor=DRbObject.new(nil, "druby://127.0.0.1:8888")
        sleep 3 # Allow the DRb Server to fully start up etc
        @current_chain=[]
        debug_info "Startup done!"
    end

    def start_clean_word
        @word_conn.close if @word_conn
        retry_count=RETRIES
        begin
            debug_info "Starting clean Word process.."
            @word_conn=Connector.new(CONN_OFFICE, 'word')
            @current_pid=@word_conn.pid
        rescue
            debug_info "Failed to start Word, retrying. #{$!}"
            sleep(1) and retry unless (retry_count-=1)<=0
            raise "Couldn't establish connection to app. #{$!}"
        end
        debug_info "New Word process pid #{@current_pid}."
        @word_conn.visible=@agent_options['visible']
    end

    def setup_for_delivery( delivery_options, preserve_chain=false )
        retry_count=RETRIES
        begin
            debug_info "Starting monitor..."
            start_clean_word
            @current_chain.clear unless preserve_chain
            @monitor.start( @word_conn.pid, @word_conn.wid, delivery_options )
        rescue
            debug_info "Failed to start monitor, retrying. #{$!}"
            sleep(1) and retry unless (retry_count-=1)<=0
            raise "#{COMPONENT}:#{VERSION}: Failed to setup for delivery. Fatal."
        end
    end

    def deliver( filename, delivery_options={} )
        debug_info "New delivery. #{filename}"
        status='error'
        exception_data=''
        chain=''
        delivery_options=DELIVERY_DEFAULTS.merge( delivery_options )
        begin
            if @monitor.exception_data
                puts "UNHANDLED CRASH!!" # the ultimate sin
                @unhandled||=0
                @unhandled+=1
                # Do something better with these later. For now, don't lose 'em!
                File.open("C:/unhandled-#{@unhandled}.doc","wb+") {|io| io.write( @current_chain.last )}
                File.open("C:/unhandled-info-#{@unhandled}.txt","wb+") {|io| io.write( @monitor.exception_data )}
                @monitor.clear_exception
            end
            if delivery_options['clean'] or not (@word_conn && @word_conn.connected?)
                @monitor.reset
                setup_for_delivery( delivery_options )
            end
            begin
                @word_conn.visible=@agent_options['visible'] # if this raises the app has died or hung
                @monitor.new_test filename
            rescue
                debug_info "Monitor reports fault in new_test, Setting up again."
                setup_for_delivery( delivery_options )
            end
            # Always keep file chains, but only send them back
            # when the filechain option is set. Uses more RAM
            # but it makes no sense to be able to set this option per
            # test.
            @current_chain << File.open( filename, "rb") {|io| io.read}
            retry_count=RETRIES
            begin
                # Open and then close the document via OLE
                @word_conn.blocking_write( filename ).close
                raise unless @monitor.running?
                status='success'
            rescue
                unless @monitor.running?
                    raise "#{COMPONENT}:#{VERSION} Too many faults, giving up." if (retry_count-=1)<=0
                    debug_info "Monitor reports fault. Delivering again."
                    setup_for_delivery( delivery_options, preserve_chain=true )
                    retry
                end
                status='fail'
            end
            @monitor.last_tick
            if @monitor.hang? #overwrites previous status
                status='hang'
                @monitor.clear_hang
                @word_conn.close rescue nil
                @word_conn=nil
            end
            if @monitor.exception_data
                status='crash' # overwrites previous status
                exception_data=@monitor.exception_data
                @monitor.clear_exception
                chain=@current_chain if delivery_options['filechain']
                @word_conn.close rescue nil
                @word_conn=nil
            end
            if delivery_options['clean'] or @current_chain.size >= delivery_options['maxchain']
                @word_conn.close rescue nil
                @word_conn=nil
            end
            debug_info "STATUS: #{status}"
            [status,exception_data,chain]
        rescue
            warn $@
            exit
        end
    end

    def destroy
        debug_info "Received destroy..."
        @word_conn.close rescue nil
        @monitor.destroy
    end

    def method_missing( meth, *args )
        debug_info "MM: #{meth}"
        @word_conn.send( meth, *args )
    end

end
