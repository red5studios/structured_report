require 'rubygems'
require 'bundler/setup'

environment = :test

Bundler.require(:default, environment)

require 'rspec'
require_relative '../lib/structured_report'