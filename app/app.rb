require 'sinatra'
require 'pry-byebug'
require File.expand_path(File.dirname(__FILE__) +'/../lib/excel_merge')
require 'tempfile'

post '/' do
  xls_template = params['template'][:tempfile]
  samples_csv = params['csv'][:tempfile]

  manifest = ExcelMerge.merge_template_with_csv(xls_template, samples_csv)
  response.headers['content_type'] = "application/vnd.ms-excel"
  Tempfile.open('manifest.xls') do |temp_manifest|
    manifest.write(temp_manifest.path)

    temp_manifest.open
    send_file(temp_manifest.path)

  end


end

