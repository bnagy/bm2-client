# Author: Ben Nagy
# Copyright: Copyright (c) Ben Nagy, 2006-2010.
# License: The MIT License
# (See README.TXT or http://www.opensource.org/licenses/mit-license.php for details.)

# The main thing this class does is overload the deliver method in the FuzzClient class to 
# do the Word specific delivery stuff.  This is the key file that would have to be rewritten
# to change fuzzing targets.
#
# In my setup, this file is invoked by a batch script that runs at system startup, and
# copies the neccessary scripts from a share, so to upgrade this code you can just change
# the shared copy and reboot all your fuzzclients.

require File.dirname(__FILE__) + '/../bm2-core/fuzz_client'
require File.dirname(__FILE__) + '/../bm2-core/connector'
require File.dirname(__FILE__) + '/conn_office'
require File.dirname(__FILE__) + '/conn_cdb'
require File.dirname(__FILE__) + '/word_delivery_agent'
require 'fileutils'
require 'rubygems'
require 'trollop'

OPTS = Trollop::options do 
    opt :debug, "Turn on debug mode", :type => :boolean
    opt :poll_interval, "Poll Interval", :type=>:integer, :default=>60
    opt :server, "Server to connect to", :type=>:string, :default=>"192.168.122.1" # for qemu
end


class WordFuzzClient < FuzzClient
    VERSION="3.5.0"

    def prepare_test_file( data )
        begin
            @test_count||=0
            filename="#{@test_count+=1}.doc"
            path=File.join(self.class.work_dir,filename)
            File.open(path, "wb+") {|io| io.write data}
            path
        rescue
            raise RuntimeError, "Fuzzclient: Couldn't create test file #{filename} : #{$!}"
        end
    end

    def clean_up( fn )
        10.times do
            begin
                FileUtils.rm_f(fn)
            rescue
                raise RuntimeError, "Fuzzclient: Failed to delete #{fn} : #{$!}"
            end
            return true unless File.exist? fn
            sleep(0.1)
        end
        return false
    end

    def deliver( test, delivery_options )
        begin
            @delivery_agent||=WordDeliveryAgent.new
            fname=prepare_test_file( test )
            status, details, chain=@delivery_agent.deliver( fname, delivery_options )
            clean_up( fname )
            [status, details, chain]
        rescue
            ["error: #{$!}",'',[]]
        end
    end

end


WordFuzzClient.setup(
    'server_ip'=>OPTS[:server],
    'work_dir'=>'R:/fuzzclient',
    'debug'=>OPTS[:debug],
    'poll_interval'=>OPTS[:poll_interval],
    'queue_name'=>'word'
)

EventMachine::run {
    EventMachine::connect(WordFuzzClient.server_ip,WordFuzzClient.server_port, WordFuzzClient)
}
puts "Event loop stopped. Shutting down."
