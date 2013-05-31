require 'sinatra'
require File.expand_path(File.dirname(__FILE__) +'/../lib/excel_merge')
require 'tempfile'

post '/' do
  xls_template = params['template'][:tempfile]
  samples_csv  = params['manifest-details'][:tempfile]
  manifest     = ExcelMerge.merge_template_with_csv(xls_template, samples_csv)


  Tempfile.open('manifest.xls') do |temp_manifest|
    manifest.write(temp_manifest)
    temp_manifest.open
    response.headers['content_type'] = "application/vnd.ms-excel"
    send_file(temp_manifest)
  end
end

