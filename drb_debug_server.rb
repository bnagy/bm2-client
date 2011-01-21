# Author: Ben Nagy
# Copyright: Copyright (c) Ben Nagy, 2006-2010.
# License: The MIT License
# (See README.TXT or http://www.opensource.org/licenses/mit-license.php for details.)

require File.dirname(__FILE__) + '/../bm2-core/connector'
require File.dirname(__FILE__) + '/conn_cdb'
require 'rubygems'
require 'trollop'
require 'drb'

OPTS=Trollop::options do
    opt :port, "Port to listen on, default 8888", :type=>:integer, :default=>8888
    opt :debug, "Debug output", :type=>:boolean
end

class DebugServer
    COMPONENT="DebugServer"
    VERSION="1.5.0"

    def start_debugger( *args )
        @this_debugger=Connector.new(CONN_CDB, *args)
        @subserver=DRb.start_service( nil, @this_debugger )
        warn "#{COMPONENT}:#{VERSION}: Started #{@this_debugger.debugger_pid} for #{args[0]['pid']}" if OPTS[:debug]
        # Return the pids so the client code can cache them
        [@this_debugger.debugger_pid, @this_debugger.target_pid, @subserver.uri]
    rescue
        warn $@
        raise $!
    end

    def close_debugger
        warn "#{COMPONENT}:#{VERSION}: Closing #{@this_debugger.debugger_pid rescue -1}" if OPTS[:debug]
        Thread.critical do
            @this_debugger.close if @this_debugger
            @this_debugger=nil
            @subserver.stop_service if @subserver
            @subserver=nil
        end
        warn "#{COMPONENT}:#{VERSION}: Closed" if OPTS[:debug]
    rescue
        warn $@
        raise $!
    end

    def destroy
        begin
            warn "#{COMPONENT}:#{VERSION}: Received destroy. Exiting." if OPTS[:debug]
            close_debugger rescue nil
        rescue
            puts $!
        ensure
            Process.exit!
        end
    end
end

DRb.start_service( "druby://127.0.0.1:#{OPTS[:port]}", DebugServer.new )
DRb.thread.join
