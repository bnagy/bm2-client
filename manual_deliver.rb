# Author: Ben Nagy
# Copyright: Copyright (c) Ben Nagy, 2006-2010.
# License: The MIT License
# (See README.TXT or http://www.opensource.org/licenses/mit-license.php for details.)

require File.dirname(__FILE__) + '/word_delivery_agent'
require 'rubygems'
require 'trollop'

OPTS = Trollop::options do 
    opt :log, "Print output to ./manualdeliver.log instead of stdout", :type => :boolean
    opt :norepair, "Open with without repair (not default for Word)", :type=> :boolean
    opt :clean, "Open a new process for each test", :type=> :boolean
    opt :invisible, "Hide Word app window", :type=>:boolean
    opt :debug, "Print debug info to stderr", :type => :boolean
    opt :grind, "Infinite loop.", :type=>:boolean
end

output=( OPTS[:log] ? File.open( "manualdeliver.log", "wb+" ) : $stdout )

delivery_options={}
delivery_options['clean']=OPTS[:clean]
delivery_options['norepair']=OPTS[:norepair]

warn "ManualDeliver: Opts: #{delivery_options.inspect}" if OPTS[:debug]

# So this has some filenames hardwired in from when I was testing it.
# That should probably be changed.
begin
    w=WordDeliveryAgent.new( 'visible'=>!(OPTS[:invisible]), 'debug'=>OPTS[:debug] )
    # Don't deliver undeleted temp files.
    ARGV.reject! {|fn| fn=~/~\$/}
    results={}
    counter=0
    cumulative_time=0
    loop do
        fname=ARGV.sample
        mark=Time.now
        output.puts "Trying #{fname}"
        status, details, dump, chain=w.deliver( fname, delivery_options )
        cumulative_time+=Time.now - mark
        output.puts "FILENAME: #{fname} STATUS: #{status} TIME: #{Time.now - mark}"
        if results[fname] and not (fname=~/TLEJQ-1392780.doc/ or fname=~/TLEJQ-1099592.doc/)
            fail "@#{counter} FUCK - UNRELIABLE. #{status} versus #{results[fname]}" unless status==results[fname]
        else
            results[fname]=status
        end
        fail "@#{counter} FUCK FALSE NEGATIVE" if fname=~/crash/ and status!='crash'
        fail "@#{counter} FUCK FALSE POSITIVE" if not fname=~/crash/ || fname=~/TLEJQ-1392780.doc/ and status=='crash'
        break unless OPTS[:grind]
        counter+=1
        output.puts "[][][]==> AVERAGE SPEED #{cumulative_time / counter}" if counter%100==0
    end
rescue Exception=>e
    warn $!
    warn e.backtrace
ensure
    output.puts "[][][]==> AVERAGE SPEED #{cumulative_time / counter}" if counter%100==0
    w.destroy
end
