require 'sinatra'
require File.expand_path(File.dirname(__FILE__) +'/../lib/excel_merge')
require 'tempfile'

post '/merge_excel/' do
  xls_template = params['template'][:tempfile]
  samples_csv  = params['manifest-details'][:tempfile]
  manifest     = ExcelMerge.merge_template_with_csv(xls_template, samples_csv)

  manifest.write(manifest.io)
  response.headers['content_type'] = "application/vnd.ms-excel"
  send_file(manifest.io)
end

