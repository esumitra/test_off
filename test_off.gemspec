Gem::Specification.new do |s|
    s.platform  =   Gem::Platform::RUBY
    s.name      =   "test_off"
    s.version   =   "0.2"
    s.author    =   "Edward Sumitra"
    s.email     =   "ed_sumitra@yahoo.com"
    s.summary   =   "A library for Behavior Driven Development of VBA code"
    s.requirements = "MS Office, rspec and cucumber to run tests"
    s.files     =   ["lib/msword.rb", "features/word_app.feature", "features/word_doc.feature", "features/data/DocWithMacros.doc", "features/data/DocWithoutMacros.doc", "features/step_definitions/word_app_steps.rb", "features/step_definitions/word_doc_steps.rb", "test/rspec/document_spec.rb", "test/rspec/word_spec.rb"]
    s.require_path  =   "lib"
    s.has_rdoc  =   false
    s.date = %q{2009-07-13}
    s.description = %q{A library for Behavior Driven Development of VBA code}
    s.homepage = %q{http://github.com/esumitra/test_off}
end
