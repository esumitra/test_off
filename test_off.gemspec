Gem::Specification.new do |s|
    s.platform  =   Gem::Platform::RUBY
    s.name      =   "test_off"
    s.version   =   "0.2"
    s.author    =   "Edward Sumitra"
    s.email     =   "ed_sumitra@yahoo.com"
    s.summary   =   "A library for Behavior Driven Development of VBA code"
    s.requirements = "MS Office, rspec and cucumber to run tests"
    s.files     =   FileList['lib/*.rb', 'features/*.*','features/data/*.*','features/step_definitions/*.*','test/**/*.*'].to_a
    s.require_path  =   "lib"
    s.has_rdoc  =   false
    s.date = %q{2009-07-13}
    s.description = %q{A library for Behavior Driven Development of VBA code}
    s.homepage = %q{http://github.com/esumitra/test_off}
end
