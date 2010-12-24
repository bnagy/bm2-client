# Author: Ben Nagy
# Copyright: Copyright (c) Ben Nagy, 2006-2010.
# License: The MIT License
# (See README.TXT or http://www.opensource.org/licenses/mit-license.php for details.)

require 'rubygems'
require 'win32/process'
require 'sys/proctable'

module ProcessKiller
    RETRY_COUNT=5

    def pids( caption )
        # This uses WMI. From my experience that might be a little
        # more fragile than ToolHelp32Snapshot, but it's much easier
        # and cleaner code.
        Sys::ProcTable.ps.to_a.select {|p|
            p.caption.upcase==caption.upcase
        }.map {|p| p.pid}
    end

    def kill_all( signal, pid_ary )
        pid_ary.each {|pid| Process.kill( signal, pid ) rescue nil }
    end

    def slay( caption )
        retry_count=RETRY_COUNT
        loop do
            kill_all 9, pids
            return if (pids=pids( caption )).empty?
            raise "#{COMPONENT}:#{VERSION}: #{__method__}( #{caption} ) exceeded retries." if (retry_count-=1) <= 0
            sleep 1
        end
    end

    def nicely_kill( caption )
        retry_count=RETRY_COUNT
        loop do
            kill_all 1, pids
            return if (pids=pids( caption )).empty?
            if (retry_count-=1) <= 0
                slay caption
                return
            end
            sleep 1
        end
    end

    module_function :nicely_kill, :slay, :pids

end
