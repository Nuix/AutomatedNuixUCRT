require 'thread'
require 'json'
require 'date'
require 'csv'
require 'fileutils'
require 'rexml/document'

# This file re-opens some standard Ruby classes and extends them with useful methods
class Numeric
    SIZE_BASE_UNIT = 1024.0
    # Converts numeric self from bytes to terabytes
    def to_tb(decimal_places)
        gb_f = self.to_f / (SIZE_BASE_UNIT ** 4)
        if decimal_places == 0
            return gb_f.to_i
        else
            return gb_f.round(decimal_places)
        end
    end
    # Converts numeric self from bytes to gigabytes
    def to_gb(decimal_places)
        gb_f = self.to_f / (SIZE_BASE_UNIT ** 3)
        if decimal_places == 0
            return gb_f.to_i
        else
            return gb_f.round(decimal_places)
        end
    end
    # Converts numeric self from bytes to megabytes
    def to_mb(decimal_places)
        gb_f = self.to_f / (SIZE_BASE_UNIT ** 2)
        if decimal_places == 0
            return gb_f.to_i
        else
            return gb_f.round(decimal_places)
        end
    end
    # Converts numeric self from bytes to kilobytes
    def to_kb(decimal_places)
        gb_f = self.to_f / (SIZE_BASE_UNIT)
        if decimal_places == 0
            return gb_f.to_i
        else
            return gb_f.round(decimal_places)
        end
    end
    # Converts numeric self to string with most sensical size
    def to_filesize(decimal_places)
        if self >= (SIZE_BASE_UNIT.to_i ** 4)
            return "#{self.to_tb(decimal_places).with_commas} TB"
        elsif self >= (SIZE_BASE_UNIT.to_i ** 3)
            return "#{self.to_gb(decimal_places).with_commas} GB"
        elsif self >= (SIZE_BASE_UNIT.to_i ** 2)
            return "#{self.to_mb(decimal_places).with_commas} MB"
        elsif self >= (SIZE_BASE_UNIT.to_i)
            return "#{self.to_kb(decimal_places).with_commas} KB"
        else
            return "#{self.with_commas} B"
        end
    end
    # Converts a number representing seconds elapsed into a string
    # 65 seconds becomes "00:01:05"
    def to_elapsed
        Time.at(self).gmtime.strftime("%H:%M:%S")
    end
end
class Integer
    # Converts an integer to a comma formatted string
    # 1337 becomes "1,337"
    def with_commas
        return self.to_s.reverse.gsub(/...(?=.)/,'\&,').reverse
    end
end
class Float
    # Converts a float to a comma formatted string
    # 1337.14159 becomes "1,337.14159"
    def with_commas
        parts = self.to_s.split('.')
        parts[0] = parts[0].to_s.reverse.gsub(/...(?=.)/,'\&,').reverse
        return parts.join(".")
    end
end

class Array
	def second
		self.length <= 1 ? nil : self[1]
	end
	def third
		self.length <= 2 ? nil : self[2]
	end
	def fourth
		self.length <= 3 ? nil : self[3]
	end
end

class Date
  def quarter
    case self.month
    when 1,2,3
      return 1
    when 4,5,6
      return 2
    when 7,8,9
      return 3
    when 10,11,12
      return 4
    end
  end
  
  def dayname
     DAYNAMES[self.wday]
  end
end

class TimeSpanFormatter
	SECOND = 1
	MINUTE = 60 * SECOND
	HOUR = 60 * MINUTE
	DAY = 24 * HOUR

	def self.format_seconds(seconds)
		# Make sure were using whole numbers
		seconds = seconds.to_i
		days = seconds / DAY
		seconds -= days * DAY

		hours = seconds / HOUR
		seconds -= hours * HOUR

		minutes = seconds / MINUTE
		seconds -= minutes * MINUTE
		days_string = ""
		if days > 0
			days_string = "#{days} Days"
		end
		hours_string = hours.to_s.rjust(2,"0")
		minutes_string = minutes.to_s.rjust(2,"0")
		seconds_string = seconds.to_s.rjust(2,"0")

		if days > 0
			return "#{days_string} #{hours_string}:#{minutes_string}:#{seconds_string}"
		else
			return "#{hours_string}:#{minutes_string}:#{seconds_string}"
		end
	end
end

class CONVERTSIZE
	def self.tokb(sizeinbytes)
		sizeinkb = sizeinbytes/1024
		return sizeinkb
	end

	def self.tomb(sizeinbytes)
		sizeinmb = sizeinbytes/1024/1024
		return sizeinmb
	end

	def self.togb(sizeinbytes)
		sizeingb = sizeinbytes/1024/1024/1024
		return sizeingb
	end
end

# settings class to get values from the settings.json file
class Settings
	def self.getDBlocation(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		dblocation = datahash['dblocation']
		return dblocation
	end

	def self.getExporttype(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		exporttype = datahash['exporttype']
		return exporttype
	end

	def self.getExportdirectory(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		exportdirectory = datahash['exportdirectory']
		return exportdirectory
	end

	def self.getExportFilename(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		exportfilename = datahash['exportfilename']
		return exportfilename
	end
	
	def self.getExportCaseInfo(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		exportcaseinfo = datahash['exportcaseinfo']
		return exportcaseinfo
	end
	
	def self.getUCRTreportinguser(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		reportinguser = datahash['ucrtreportinguser']
		return reportinguser
	end

	def self.getReportfrequency(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		reportrunfrequency = datahash['report-run-frequency']
		return reportrunfrequency
	end

	def self.getUpgradeValue(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		upgradevalue = datahash['upgradecases']
		return upgradevalue
	end

	def	self.getCleanupDatabaseValue(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		cleanupdatabase = datahash['cleanupdatabase']
		return cleanupdatabase
	end

	def	self.getCleanupDatabaseGUIDs(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		cleanupcaseguids = datahash['cleanupcaseguids']
		return cleanupcaseguids
	end	

	def self.getIncludeDateRangeValue(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		includedaterangevalue = datahash['includedaterange']
		return includedaterangevalue
	end

	def self.getIncludeAnnotationsValue(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		includeannotations = datahash['includeannotations']
		return includeannotations
	end

	def self.getIgnoreCaseHistoryValue(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		ignorecasehistory = datahash['ignorecasehistory']
		return ignorecasehistory
	end
	
	def self.getCSVExportFields(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		csvexportfields = datahash['exportfields']
		return csvexportfields
	end

	def self.getShowSizeInValue(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		showsizein = datahash['showsizein']
		return showsizein
	end
	
	def self.getSearchTermValue(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		searchterm = datahash['searchterm']
		return searchterm
	end

	def self.getSearchTermFileValue(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		searchtermfile = datahash['searchtermfile']
		return searchtermfile
	end
	
	def self.getVersionInfo(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		versioninfo = datahash['nuixversionmapping']
		return versioninfo
	end

	def self.getDecimalPointAccuracy(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		decimal_places = datahash['decimalpointaccuracy']
		return decimal_places
	end	
	def self.getCleanupFilesValue(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		cleanupfiles = datahash['cleanupfiles']
		return cleanupfiles
	end
	def self.getCleanupFileRange(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		cleanup_filerange = datahash['cleanupfilerange']
		return cleanup_filerange
	end
	def self.getCleanupDirectories(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		cleanupdirectories = datahash['cleanupdirectories']
		return cleanupdirectories
	end	
	def self.getCleanupFileType(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		cleanupfiletypes = datahash['cleanupfilestype']
		return cleanupfiletypes
	end	
	def self.getNuixExportSearchResults(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		nuixexportsearchresults = datahash['nuixexportsearchresults']
		return nuixexportsearchresults
	end	
	def self.getNuixExportType(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		datahash = JSON.parse(settingsfile)
		nuixexporttype = datahash['nuixexporttype']
		return nuixexporttype
	end	


end
class ItemStatGenerator	
	attr_accessor :stats
	attr_accessor :digests

	def initialize
		@digests = {}
		@stats = Hash.new{|h,k| h[k] = Hash.new{|h2,k2| h2[k2] = {:count => 0, :filesize => 0} } }
	end

	def collect(items)
		items.each do |item|
			md5 = item.getDigests.getMd5
			is_original = true
			if !md5.nil? && !md5.strip.empty?
				if @digests.has_key?(md5)
					is_original = false
				else
					is_original = true
					@digests[md5] = true
				end
			end
			kind = item.getKind.getName
			filesize = item.getFileSize
			status = is_original ? :original : :duplicate

			if !filesize.nil?
				@stats[kind][status][:filesize] += filesize
				@stats["TOTAL"][status][:filesize] += filesize
			end
			
			@stats[kind][status][:count] += 1
			@stats["TOTAL"][status][:count] += 1
		end
	end
end
# UCRT class to get the specific statistical information from the nuix case
class UCRT
	# get all the cases from from the folders that are specified in the settings.json file
	def self.getCases(ucrtsettings)
		settingsfile = File.read("#{ucrtsettings}")
		z = []
		z << ENV["USERDOMAIN"]
		z << ENV["USERNAME"]
		loggedinuser = z.join("\\")
		nuixcases = Array.new
		datahash = JSON.parse(settingsfile)
		caseslocations = datahash['caseslocations']
		puts "Cases Locations = #{caseslocations}"
		caseslocationsarray = caseslocations.split(",")
		# loop over the cases locations specified in the caselocations in the settings.json file
		caseslocationsarray.each do |caseslocation|
			# can the loggedinuser see the caselocation
			if java.io.File.new(caseslocation).exists
				caselocation = caseslocation.gsub("/\/","/")
				casefiles = "#{caseslocation}/**/case.fbi2"
				Dir.glob(casefiles).each do |caselocation|
					# add the caselocation nuixcases array
					nuixcases << caselocation
				end
			else
				puts "#{loggedinuser} cannot see #{caseslocation}"
			end
		end
		puts "Nuix cases = #{nuixcases}"
		# return the nuixcases array
		return nuixcases
	end

	# get details of the case on disc - these values are not values that the case has to be opened to 
	# obtain - these are values that are on disc or in files in the case directory
	def self.getCaseDetails(caselocation)
		begin
			case_investigator = ''
			nuix_version = ''
			case_type = ''
			case_name = ''
			creation_date = ''
			modified_date = ''
			case_guid = ''
			evidence_name = ''
			evidence_locations = ''
			evidence_description = ''
			worker_temp_dir = ''
			worker_count = ''
			worker_memory = ''
			broker_memory = ''
			is_locked = ''
			locked_by = ''
			lock_machine = ''
			lock_date = ''
			lock_product = ''
			case_details = {}
			created_nuix_version = ''
			saved_nuix_version = ''
			casefiles = "#{caselocation}/**/case.fbi2"
			# look over all the files in the caselocation and look for a file called case.fbi2 (which is an xml file)
			Dir.glob(casefiles).each do |caselocation|
				xml_file = File.new(caselocation)
				modified_date = File.mtime(xml_file)
				# parse the xml file 
				xml_doc = REXML::Document.new(xml_file)
				# get the investigator from the xml element
				xml_doc.elements.each("case/metadata/investigator") do |investigator|
					case_investigator = investigator.text
				end
				# get the created-by-product from the created-by-product xml element
				xml_doc.elements.each("case/metadata/created-by-product") do |version_node|
					created_nuix_version = version_node.attributes["version"]
				end
				# get the saved-by-product from the saved-by-product xml element
				xml_doc.elements.each("case/metadata/saved-by-product") do |version_node|
					saved_nuix_version = version_node.attributes["version"]
				end
				# get the caseType from the caseType xml element
				xml_doc.elements.each("case/metadata/caseType") do |casetype|
					case_type = casetype.text
				end
				# get the name from the name xml element
				xml_doc.elements.each("case/metadata/name") do |name|
					case_name = name.text
				end
				# get the creation-date from the creation-date xml element
				xml_doc.elements.each("case/metadata/creation-date") do |creationdate|
					creation_date = creationdate.text
				end
				# get the guid from the guid xml element
				xml_doc.elements.each("case/metadata/guid") do |caseguid|
					case_guid = caseguid.text
				end
			end

			# determine what the version of the case is - since there are so many different potential versions
			# determine what the max version is from all the different types
			puts "Created Nuix Version - #{created_nuix_version}"
			puts "Save Nuix Version - #{saved_nuix_version}"
			if created_nuix_version == saved_nuix_version
				nuix_version = created_nuix_version
			elsif created_nuix_version != '' && saved_nuix_version == ''
				nuix_version = created_nuix_version
			elsif saved_nuix_version != '' && created_nuix_version == ''
				nuix_version = saved_nuix_version
			elsif saved_nuix_version != '' && created_nuix_version != ''
				created_nuix_ver_array = created_nuix_version.split(".")
				created_array_count = created_nuix_ver_array.count
				saved_nuix_ver_array = saved_nuix_version.split(".")
				saved_array_count = saved_nuix_ver_array.count
				firstsavedelement = saved_nuix_ver_array.first
				firstcreatedelement = created_nuix_ver_array.first
				if firstsavedelement > firstcreatedelement
					nuix_version = saved_nuix_version
				elsif firstcreatedelement > firstsavedelement
					nuix_version = created_nuix_version
				elsif firstsavedelement == firstcreatedelement
					secondsavedelement = saved_nuix_ver_array.second
					secondcreatedelement = created_nuix_ver_array.second
					if secondsavedelement > secondcreatedelement
						nuix_version = saved_nuix_version
					elsif secondcreatedelement > secondsavedelement
						nuix_version = created_nuix_version
					elsif secondsavedelement == secondcreatedelement
						thirdsavedelement = saved_nuix_ver_array.third
						thirdcreatedelement = created_nuix_ver_array.third
						if thirdsavedelement > thirdcreatedelement
							nuix_version = saved_nuix_version
						elsif thirdcreatedelement > thirdsavedelement
							nuix_version = created_nuix_version
						elsif thirdsavedelement == thirdcreatedelement
							if created_array_count >= 3
								fourthcreatedelement = created_nuix_ver_array.fourth
							end
							if saved_array_count >= 3
								fourthsavedelement = saved_nuix_ver_array.fourth
							end
							if fourthcreatedelement != '' && fourthsavedelement == ''
								nuix_version = created_nuix_version
							elsif fourthsavedelement != '' && fourthcreatedelement == ''
								nuix_version = saved_nuix_version
							elsif fourthsavedelement > fourthcreatedelement
								nuix_version = saved_nuix_version
							elsif fourthcreatedelement > fourthsavedelement
								nuix_version = created_nuix_element
							end
						end
					end
				end 
			end
			puts "Nuix Case version is - #{nuix_version}"
			
			# check for the existence of a case.lock file
			lockfile = "#{caselocation}/case.lock"
			# if a lockfile exists
			if File.exist?(lockfile)
				# get the lockfile modified time
				lock_date = File.mtime(lockfile)
				# set the value of is_locked to true
				is_locked = 'true'
				# get the file case.lock.properties
				lockpropfile = File.new("#{caselocation}/case.lock.properties")
				# open the case.lock.properties file and loop through each line of the file
				File.open(lockpropfile).each do |lockline|
					# strip the line
					lockline = lockline.strip
					# get the product
					if lockline.include? "product"
						aslockproduct = lockline.split("=")
						# get the lock_product
						lock_product = aslockproduct.second
					# get the host
					elsif lockline.include? "host"
						aslockmachine = lockline.split("=")
						# get the lock machine
						lock_machine = aslockmachine.second
					# get the user
					elsif lockline.include? "user"
						aslockuser = lockline.split("=")
						# get the locked by person
						locked_by = aslockuser.second
					end
				end

			end
			# find the parallelprocessingsettings.properties file
			parallelprocessingfiles = "#{caselocation}/**/parallelprocessingsettings.properties"
			# loop over the directory until the file is found
			Dir.glob(parallelprocessingfiles).each do |ppfile|
				line_num=0
				pp_filename = File.basename(ppfile)
				# open the file and loop through each line
				File.open(ppfile).each do |line|
					# get the workerTempDirectory
					if line.include? "workerTempDirectory"
						asWorkerTemp = line.split("=")
						worker_temp_dir = asWorkerTemp.second
					# get the workerCount
					elsif line.include? "workerCount"
						asWorkerCount = line.split("=")
						worker_count = asWorkerCount.second
					# get the brokerMemory
					elsif line.include? "brokerMemory"
						asBrokerMemory = line.split("=")
						broker_memory = asBrokerMemory.second
					# get the workerMemory
					elsif line.include? "workerMemory"
						asWorkerMemory = line.split("=")
						worker_memory = asWorkerMemory.second
					end
				end
			end

			# find the evidence xml files
			evidencexmlfiles= "#{caselocation}/**/*.xml"
			# recurse through the directories to fine the evidence xmlfiles
			Dir.glob(evidencexmlfiles).each do |evidencefiles|
				evidencexml_file = File.new(evidencefiles)
				evidencexml_filename = File.basename(evidencefiles)
				# if the evidence file is the summart_report.xml
				if evidencexml_filename != "summary_report.xml"
					evidencexml_doc = REXML::Document.new(evidencexml_file)
					evidencexml_doc.elements.each("evidence/name") do |evidencename|
						evidence_name += evidencename.text + ';'
					end
					evidencexml_doc.elements.each("evidence/data-roots") do |dataroots|
						dataroots.elements.each("data-root/file") do |dataroot|
		#					puts dataroot.attributes["location"]
							evidence_locations.concat(dataroot.attributes["location"])
							evidence_locations.concat(';')
						end
					end
					evidencexml_doc.elements.each("evidence/description") do |evidencedescription|
						if evidencedescription.text.to_s != ''
							evidence_description += evidencedescription.text + ';'
						end
					end
				end
			end
			# set the case details map with all the case info collected
			case_details = {"case_name" => case_name, "case_guid" => case_guid, "case_type" => case_type, "creation_date" => creation_date, "investigator" => case_investigator, "nuix_version" => nuix_version,"workerTemp" => worker_temp_dir, "workerCount" => worker_count, "brokerMemory" => broker_memory, "workerMemory" => worker_memory,"evidence_name" => evidence_name,"evidence_locations" => evidence_locations,"evidence_description" => evidence_description, "modified_date" => modified_date, "is_locked" => is_locked, "locked_by" => locked_by, "lock_machine" => lock_machine, "lock_date" => lock_date, "lock_product" => lock_product}
			# return the case details
			return case_details 
		rescue StandardError => msg
			puts "Error getting Case Details - #{msg}"
		end
	end

	# get case statistics - these values obtained from opening the case and using the API
	def self.getCaseStats(nuix_case, case_guid, showsizein, decimal_places, db)
		begin
			# get the case guid
			case_guid = nuix_case.getGuid
			# get the description of the case
			case_description = nuix_case.getDescription

			# remove the - from the case guid
			case_guid = case_guid.tr("-","")
			# remove any trailing spaces from the case name
			case_name = nuix_case.getName().strip
			# get the count of all items in the case
			total_item_count = nuix_case.count('')
			
			# get the total file size of all items in the case
			total_file_size = nuix_case.getStatistics.getFileSize('')
			# get the total audit size of all items in the case
			total_audit_size = nuix_case.getStatistics.getAuditSize('')
			# convert the size from bytes to the value of the showsizein value
			case showsizein
				when "byte"
					total_file_size = total_file_size
					total_audit_size = total_audit_size
				when "kb"
					total_file_size = total_file_size.to_kb(decimal_places)
					total_audit_size = total_audit_size.to_kb(decimal_places)
				when "mb"
					total_file_size = total_file_size.to_mb(decimal_places)
					total_audit_size = total_audit_size.to_mb(decimal_places)
				when "gb"
					total_file_size = total_file_size.to_gb(decimal_places)
					total_audit_size = total_audit_size.to_gb(decimal_places)
				else
					total_file_size = total_file_size
					total_audit_size = total_audit_size
			end

			# get the value of the isCompount
			compound_case_contains = ''
			is_compound = nuix_case.isCompound()
			# if the case is compound then get all the child cases
			if "#{is_compound}"== 'true'
				child_cases = nuix_case.getChildCases
				child_cases.each do |cases|
					compound_case_contains = compound_case_contains + cases.getName() + ';'
				end
			end
			# get all case users
			all_case_users = ''
			case_users = nuix_case.getAllUsers()
			# loop over all case users
			case_users.each do |case_user|
				all_case_users = all_case_users + "#{case_user}" + ';'
			end
			# all the data collected to this point to the database
			case_users_data = [total_item_count, total_file_size, total_audit_size, all_case_users, case_description, is_compound.to_s, compound_case_contains, 5]
			db.update("UPDATE NuixReportingInfo SET TotalItemCount = ?, CaseFileSize = ?, CaseAuditSize = ?, CaseUsers = ?, CaseDescription = ?, IsCompound = ?, CasesContained = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", case_users_data)
			# get the batch load details from the case
			all_batch_load_info = ''
			batch_loads = nuix_case.getBatchLoads
			batch_loads.each do |load|
				batch_guid = "#{load.getBatchId}"
				batch_loaddate = "#{load.getLoaded}"
				batch_load_query = "batch-load-guid:#{batch_guid}"
				batch_load_count = nuix_case.count(batch_load_query)
				batch_load_file_size = nuix_case.getStatistics.getFileSize(batch_load_query)
				batch_load_audit_size = nuix_case.getStatistics.getAuditSize(batch_load_query)
				case showsizein
					when "byte"
						batch_load_file_size = batch_load_file_size
						batch_load_audit_size = batch_load_audit_size 
					when "kb"
						batch_load_file_size = batch_load_file_size.to_kb(decimal_places)
						batch_load_audit_size = batch_load_audit_size.to_kb(decimal_places)
					when "mb"
						batch_load_file_size = batch_load_file_size.to_mb(decimal_places)
						batch_load_audit_size = batch_load_audit_size.to_mb(decimal_places)
					when "gb"
						batch_load_file_size = batch_load_file_size.to_gb(decimal_places)
						batch_load_audit_size = batch_load_audit_size.to_gb(decimal_places)
					else
						batch_load_file_size = batch_load_file_size
						batch_load_audit_size = batch_load_audit_size 
				end
				
				all_batch_load_info = all_batch_load_info + "#{batch_loaddate}" + '::' + "#{batch_load_count}".to_s + '::' + "#{batch_load_file_size}".to_s + '::' + "#{batch_load_audit_size}".to_s + ";"
			end
			# update the data base with the batch load details
			case_batchload_data = [all_batch_load_info, 10]
			db.update("UPDATE NuixReportingInfo SET BatchLoadInfo = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", case_batchload_data)
		rescue StandardError => msg
			puts "Error in getCaseStats + #{msg}"
		ensure
		
		end
	end

	# get case sizes - these values obtained from opening the case and using the API to get the kinds, types, and by custodian
	def self.getCaseSize(nuix_case, case_guid, showsizein, decimal_places, db)
		begin
			# get custodian information
			puts "Getting Custodian information - #{DateTime.now}"
			custodian_count = 0
			all_custodians_info = ''
			# loop over all custodians
			custodian_names = nuix_case.getAllCustodians
			custodian_names.each do |custodian_name|
				custodian_count += 1
				search_criteria = "custodian:\"#{custodian_name}\""
				custodian_name_count = nuix_case.count("#{search_criteria}")
				all_custodians_info = all_custodians_info + "#{custodian_name}" + '::' + "#{custodian_name_count}".to_s + ";"
			end
			# update the database with the custodian information
			case_custodians_data = [all_custodians_info, custodian_count, 15]
			db.update("UPDATE NuixReportingInfo SET Custodians = ?, CustodianCount = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", case_custodians_data)
			# get the date range information - earliest items and latest items information
			puts "Getting Date Range - #{DateTime.now}"
			daterange = nuix_case.getStatistics.getCaseDateRange
			oldest = daterange.getEarliest
			newest = daterange.getLatest
			oldest={"defaultFields"=>"item-date","order"=>"item-date ASC","limit"=>1}
			newest={"defaultFields"=>"item-date","order"=>"item-date DESC","limit"=>1}
			oldestItemDate = nuix_case.search("flag:top_level item-date:*", oldest).first().getDate().to_s
			newestItemDate = nuix_case.search("flag:top_level item-date:*", newest).first().getDate().to_s
			oldestDate = Date.parse(oldestItemDate.to_s.split("T").first)
			newestDate = Date.parse(newestItemDate.to_s.split("T").first)
			# update the database with the earliest items and latest item
			case_date = [oldestDate, newestDate, 20]
			db.update("UPDATE NuixReportingInfo SET OldestItem = ?, NewestItem = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", case_date)

			# get the language information
			puts "Getting Languages - #{DateTime.now}"
			lang_detail_count = ''
			languages = nuix_case.getLanguages
			# loop over each language
			languages.each do |lang|
				lang_count = nuix_case.count("lang:#{lang}")
				lang_size = nuix_case.getStatistics.getFileSize("lang:#{lang}")
				case showsizein
					when "byte"
						lang_size = lang_size
					when "kb"
						lang_size = lang_size.to_kb(decimal_places)
					when "mb"
						lang_size = lang_size.to_mb(decimal_places)
					when "gb"
						lang_size = lang_size.to_gb(decimal_places)
					else
						lang_size = lang_size
				end
				lang_detail_count = lang_detail_count + lang + "::" + lang_count.to_s + "::" + lang_size.to_s + ";"
			end
			# add the language information to the database
			case_lang_details = [lang_detail_count, 22]
			db.update("UPDATE NuixReportingInfo SET Languages = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", case_lang_details)

			def self.sum_filesize(items)
				return items.map{|i| i.getFileSize || 0 }.reduce(0,:+)
			end

			# get the kind information for originals and duplicates and counts
			puts "Getting Kinds - #{DateTime.now}"
			original_count = 0
			original_filesize = 0
			email_original_count = 0
			email_original_filesize = 0
			calendar_original_count = 0
			calendar_original_filesize = 0
			contact_original_count = 0
			contact_original_filesize = 0
			container_original_count = 0
			container_original_filesize = 0
			image_original_count = 0
			image_original_filesize = 0
			document_original_count = 0
			document_original_filesize = 0
			spreadsheet_original_count = 0
			spreadsheet_original_filesize = 0
			drawing_original_count = 0
			drawing_original_filesize = 0
			presentation_original_count = 0
			presentation_original_filesize = 0
			database_original_count = 0
			database_original_filesize = 0
			other_document_original_count = 0
			other_document_original_filesize = 0
			no_data_original_count = 0
			no_data_original_filesize = 0
			unrecognised_original_count= 0
			unrecognised_original_filesize = 0 
			system_original_count = 0
			system_original_filesize = 0
			multimedia_original_count = 0
			multimedia_original_filesize = 0
			log_original_count = 0
			log_original_filesize = 0
			chat_conversation_original_count = 0
			chat_conversation_original_filesize = 0
			chat_message_original_count = 0
			chat_message_original_filesize = 0
			duplicate_count = 0
			duplicate_filesize = 0
			email_duplicate_count = 0
			email_duplicate_filesize = 0
			calendar_duplicate_count = 0
			calendar_duplicate_filesize = 0
			contact_duplicate_count = 0
			contact_duplicate_filesize = 0
			container_duplicate_count = 0
			container_duplicate_filesize = 0
			image_duplicate_count = 0
			image_duplicate_filesize = 0
			document_duplicate_count = 0
			document_duplicate_filesize = 0
			spreadsheet_duplicate_count = 0
			spreadsheet_duplicate_filesize = 0
			drawing_duplicate_count = 0
			drawing_duplicate_filesize = 0
			presentation_duplicate_count = 0
			presentation_duplicate_filesize = 0
			database_duplicate_count = 0
			database_duplicate_filesize = 0
			other_document_duplicate_count = 0
			other_document_duplicate_filesize = 0
			no_data_duplicate_count = 0
			no_data_duplicate_filesize = 0
			unrecognised_duplicate_count= 0
			unrecognised_duplicate_filesize = 0 
			system_duplicate_count = 0
			system_duplicate_filesize = 0
			multimedia_duplicate_count = 0
			multimedia_duplicate_filesize = 0
			log_duplicate_count = 0
			log_duplicate_filesize = 0
			chat_conversation_duplicate_count = 0
			chat_conversation_duplicate_filesize = 0
			chat_message_duplicate_count = 0
			chat_message_duplicate_filesize = 0

			# create a new ItemStatGenerator
			isg = ItemStatGenerator.new
			# search for all items in a case
			isg.collect(nuix_case.search(""))

			#loop over each of the item status
			[:original,:duplicate].each do |status|
				puts "Total #{status} Count: #{isg.stats["TOTAL"][status][:count]}"
				puts "Total #{status} File Size: #{isg.stats["TOTAL"][status][:filesize]}"
			end

			# loop over each kind and group 
			isg.stats.each do |kind,status_grouped|
				status_grouped.each do |status,stats|
					stats.each do |stat,value|
						# for each kind store all the original counts and file sizes
						if status.to_s == "original"
							case kind.to_s
								when "TOTAL"
									case stat.to_s
										when "count"
											original_count += value
										when "filesize"
											original_filesize += value
										else
									end
								when "email"
									puts "Email Original"
									case stat.to_s
										when "count"
											email_original_count += value
											puts "Email original count - #{email_original_count}"
										when "filesize"
											email_original_filesize += value
										else
									end
								when "calendar"
									case stat.to_s
										when "count"
											calendar_original_count += value
										when "filesize"
											calendar_original_filesize += value
										else
									end
								when "contact"
									case stat.to_s
										when "count"
											contact_original_count += value
										when "filesize"
											contact_original_filesize += value
										else
									end
								when "container"
									case stat.to_s
										when "count"
											container_original_count += value
										when "filesize"
											container_original_filesize += value
										else
									end
								when "image"
									case stat.to_s
										when "count"
											image_original_count += value
										when "filesize"
											image_original_filesize += value
										else
									end
								when "document"
									case stat.to_s
										when "count"
											document_original_count += value
										when "filesize"
											document_original_filesize += value
										else
									end
								when "spreadsheet"
									case stat.to_s
										when "count"
											spreadsheet_original_count += value
										when "filesize"
											spreadsheet_original_filesize += value
										else
									end
								when "drawing"
									case stat.to_s
										when "count"
											drawing_original_count += value
										when "filesize"
											drawing_original_filesize += value
										else
									end
								when "presentation"
									case stat.to_s
										when "count"
											presentation_original_count += value
										when "filesize"
											presentation_original_filesize += value
										else
									end
								when "database"
									case stat.to_s
										when "count"
											database_original_count += value
										when "filesize"
											database_original_filesize += value
										else
									end
								when "other-document"
									case stat.to_s
										when "count"
											other_document_original_count += value
										when "filesize"
											other_document_original_filesize += value
										else
									end
								when "no-data"
									case stat.to_s
										when "count"
											no_data_original_count += value
										when "filesize"
											no_data_original_filesize += value
										else
									end
								when "unrecognised"
									case stat.to_s
										when "count"
											unrecognised_original_count += value
										when "filesize"
											unrecognised_original_filesize += value
										else
									end
								when "system"
									case stat.to_s
										when "count"
											system_original_count += value
										when "filesize"
											system_original_filesize += value
										else
									end
								when "multimedia"
									case stat.to_s
										when "count"
											multimedia_original_count += value
										when "filesize"
											multimedia_original_filesize += value
										else
									end
								when "log"
									case stat.to_s
										when "count"
											log_original_count += value
										when "filesize"
											log_original_filesize += value
										else
									end
								when "chat-conversation"
										case stat.to_s
										when "count"
											chat_conversation_original_count += value
										when "filesize"
											chat_conversation_original_filesize += value
										else
									end
								when "chat-message"
									case stat.to_s
										when "count"
											chat_message_original_count += value
										when "filesize"
											chat_message_original_filesize += value
										else
									end
								else
							end
						# get all the duplicate counts and file sizes
						elsif status.to_s == "duplicate"
							case kind.to_s
								when "TOTAL"
									case stat.to_s
										when "count"
											duplicate_count += value
										when "filesize"
											duplicate_filesize += value
										else
									end
								when "email"
									case stat.to_s
										when "count"
											email_duplicate_count += value
										when "filesize"
											email_duplicate_filesize += value
										else
									end
								when "calendar"
									case stat.to_s
										when "count"
											calendar_duplicate_count += value
										when "filesize"
											calendar_duplicate_filesize += value
										else
									end
								when "contact"
									case stat.to_s
										when "count"
											contact_duplicate_count += value
										when "filesize"
											contact_duplicate_filesize += value
										else
									end
								when "container"
									case stat.to_s
										when "count"
											container_duplicate_count += value
										when "filesize"
											container_duplicate_filesize += value
										else
									end
								when "image"
									case stat.to_s
										when "count"
											image_duplicate_count += value
										when "filesize"
											image_duplicate_filesize += value
										else
									end
								when "document"
									case stat.to_s
										when "count"
											document_duplicate_count += value
										when "filesize"
											document_duplicate_filesize += value
										else
									end
								when "spreadsheet"
									case stat.to_s
										when "count"
											spreadsheet_duplicate_count += value
										when "filesize"
											spreadsheet_duplicate_filesize += value
										else
									end
								when "drawing"
									case stat.to_s
										when "count"
											drawing_duplicate_count += value
										when "filesize"
											drawing_duplicate_filesize += value
										else
									end
								when "presentation"
									case stat.to_s
										when "count"
											presentation_duplicate_count += value
										when "filesize"
											presentation_duplicate_filesize += value
										else
									end
								when "database"
									case stat.to_s
										when "count"
											database_duplicate_count += value
										when "filesize"
											database_duplicate_filesize += value
										else
									end
								when "other-document"
									case stat.to_s
										when "count"
											other_document_duplicate_count += value
										when "filesize"
											other_document_duplicate_filesize += value
										else
									end
								when "no-data"
									case stat.to_s
										when "count"
											no_data_duplicate_count += value
										when "filesize"
											no_data_duplicate_filesize += value
										else
									end
								when "unrecognised"
									case stat.to_s
										when "count"
											unrecognised_duplicate_count += value
										when "filesize"
											unrecognised_duplicate_filesize += value
										else
									end
								when "system"
									case stat.to_s
										when "count"
											system_duplicate_count += value
										when "filesize"
											system_duplicate_filesize += value
										else
									end
								when "multimedia"
									case stat.to_s
										when "count"
											multimedia_duplicate_count += value
										when "filesize"
											multimedia_duplicate_filesize += value
										else
									end
								when "log"
									case stat.to_s
										when "count"
											log_duplicate_count += value
										when "filesize"
											log_duplicate_filesize += value
										else
									end
								when "chat-conversation"
										case stat.to_s
										when "count"
											chat_conversation_duplicate_count += value
										when "filesize"
											chat_conversation_duplicate_filesize += value
										else
									end
								when "chat-message"
									case stat.to_s
										when "count"
											chat_message_duplicate_count += value
										when "filesize"
											chat_message_duplicate_filesize += value
										else
									end
								else
							end
						end
					end
				end
			end
			# store all the original and duplicate counts
			total_count = original_count + duplicate_count
			email_total_count = email_original_count + email_duplicate_count
			calendar_total_count = calendar_original_count + calendar_duplicate_count
			contact_total_count = contact_original_count + contact_duplicate_count
			document_total_count = document_original_count + document_duplicate_count
			spreadsheet_total_count = spreadsheet_original_count + spreadsheet_duplicate_count
			presentation_total_count = presentation_original_count + presentation_duplicate_count
			image_total_count = image_original_count + image_duplicate_count
			drawing_total_count = drawing_original_count + drawing_duplicate_count
			other_document_total_count = other_document_original_count + other_document_duplicate_count
			multimedia_total_count = multimedia_original_count + multimedia_duplicate_count
			database_total_count = database_original_count + database_duplicate_count
			container_total_count = container_original_count + container_duplicate_count
			system_total_count = system_original_count + system_duplicate_count
			no_data_total_count = no_data_original_count + no_data_duplicate_count
			unrecognised_total_count = unrecognised_original_count + unrecognised_duplicate_count
			log_total_count = log_original_count + log_duplicate_count
			chat_conversation_total_count = chat_conversation_original_count + chat_conversation_duplicate_count
			chat_message_total_count = chat_message_original_count + chat_message_duplicate_count

			# store all the original and dupliate filesizes
			total_filesize = original_filesize + duplicate_filesize
			email_total_filesize = email_original_filesize + email_duplicate_filesize
			calendar_total_filesize = calendar_original_filesize + calendar_duplicate_filesize
			contact_total_filesize = contact_original_filesize + contact_duplicate_filesize
			document_total_filesize = document_original_filesize + document_duplicate_filesize
			spreadsheet_total_filesize = spreadsheet_original_filesize + spreadsheet_duplicate_filesize
			presentation_total_filesize = presentation_original_filesize + presentation_duplicate_filesize
			image_total_filesize = image_original_filesize + image_duplicate_filesize
			drawing_total_filesize = drawing_original_filesize + drawing_duplicate_filesize
			other_document_total_filesize = other_document_original_filesize + other_document_duplicate_filesize
			multimedia_total_filesize = multimedia_original_filesize + multimedia_duplicate_filesize
			database_total_filesize = database_original_filesize + database_duplicate_filesize
			container_total_filesize = container_original_filesize + container_duplicate_filesize
			system_total_filesize = system_original_filesize + system_duplicate_filesize
			no_data_total_filesize = no_data_original_filesize + no_data_duplicate_filesize
			unrecognised_total_filesize = unrecognised_original_filesize + unrecognised_duplicate_filesize
			log_total_filesize = log_original_filesize + log_duplicate_filesize
			chat_conversation_total_filesize = chat_conversation_original_filesize + chat_conversation_duplicate_filesize
			chat_message_total_filesize = chat_message_original_filesize + chat_message_duplicate_filesize
	
			# convert all sizes to the requested showsizein value
			puts "Getting Convert to #{showsizein}"
			case showsizein
				when "byte"
					total_filesize = total_filesize
					email_total_filesize = email_total_filesize
					calendar_total_filesize = calendar_total_filesize 
					contact_total_filesize = contact_total_filesize
					document_total_filesize = document_total_filesize
					spreadsheet_total_filesize = spreadsheet_total_filesize 
					presentation_total_filesize = presentation_total_filesize  
					image_total_filesize = image_total_filesize 
					drawing_total_filesize = drawing_total_filesize 
					other_document_total_filesize = other_document_total_filesize 
					multimedia_total_filesize = multimedia_total_filesize 
					database_total_filesize = database_total_filesize 
					container_total_filesize = container_total_filesize 
					system_total_filesize = system_total_filesize 
					no_data_total_filesize = no_data_total_filesize 
					unrecognised_total_filesize = unrecognised_total_filesize 
					log_total_filesize = log_total_filesize 
					original_filesize = original_filesize 
					email_original_filesize = email_original_filesize 
					calendar_original_filesize = calendar_original_filesize 
					contact_original_filesize = contact_original_filesize 
					document_original_filesize = document_original_filesize 
					spreadsheet_original_filesize = spreadsheet_original_filesize 
					presentation_original_filesize = presentation_original_filesize 
					image_original_filesize = image_original_filesize 
					drawing_original_filesize = drawing_original_filesize 
					other_document_original_filesize = other_document_original_filesize 
					multimedia_original_filesize = multimedia_original_filesize 
					database_original_filesize = database_original_filesize 
					container_original_filesize = container_original_filesize 
					system_original_filesize = system_original_filesize 
					no_data_original_filesize = no_data_original_filesize 
					unrecognised_original_filesize = unrecognised_original_filesize 
					log_original_filesize = log_original_filesize 
					duplicate_filesize = duplicate_filesize 
					email_duplicate_filesize = email_duplicate_filesize 
					calendar_duplicate_filesize = calendar_duplicate_filesize 
					contact_duplicate_filesize = contact_duplicate_filesize
					document_duplicate_filesize = document_duplicate_filesize 
					spreadsheet_duplicate_filesize = spreadsheet_duplicate_filesize 
					presentation_duplicate_filesize = presentation_duplicate_filesize 
					image_duplicate_filesize = image_duplicate_filesize 
					drawing_duplicate_filesize = drawing_duplicate_filesize 
					other_document_duplicate_filesize = other_document_duplicate_filesize 
					multimedia_duplicate_filesize = multimedia_duplicate_filesize 
					database_duplicate_filesize = database_duplicate_filesize 
					container_duplicate_filesize = container_duplicate_filesize 
					system_duplicate_filesize = system_duplicate_filesize 
					no_data_duplicate_filesize = no_data_duplicate_filesize 
					unrecognised_duplicate_filesize = unrecognised_duplicate_filesize 
					log_duplicate_filesize = log_duplicate_filesize 
				when "kb"
					total_filesize = total_filesize.to_kb(decimal_places)
					email_total_filesize = email_total_filesize.to_kb(decimal_places)
					calendar_total_filesize = calendar_total_filesize.to_kb(decimal_places)
					contact_total_filesize = contact_total_filesize.to_kb(decimal_places)
					document_total_filesize = document_total_filesize.to_kb(decimal_places)
					spreadsheet_total_filesize = spreadsheet_total_filesize.to_kb(decimal_places)
					presentation_total_filesize = presentation_total_filesize.to_kb(decimal_places)
					image_total_filesize = image_total_filesize.to_kb(decimal_places)
					drawing_total_filesize = drawing_total_filesize.to_kb(decimal_places)
					other_document_total_filesize = other_document_total_filesize.to_kb(decimal_places)
					multimedia_total_filesize = multimedia_total_filesize.to_kb(decimal_places)
					database_total_filesize = database_total_filesize.to_kb(decimal_places)
					container_total_filesize = container_total_filesize.to_kb(decimal_places)
					system_total_filesize = system_total_filesize.to_kb(decimal_places)
					no_data_total_filesize = no_data_total_filesize.to_kb(decimal_places)
					unrecognised_total_filesize = unrecognised_total_filesize.to_kb(decimal_places)
					log_total_filesize = log_total_filesize.to_kb(decimal_places)
					original_filesize = original_filesize.to_kb(decimal_places)
					email_original_filesize = email_original_filesize.to_kb(decimal_places)
					calendar_original_filesize = calendar_original_filesize.to_kb(decimal_places)
					contact_original_filesize = contact_original_filesize.to_kb(decimal_places)
					document_original_filesize = document_original_filesize.to_kb(decimal_places)
					spreadsheet_original_filesize = spreadsheet_original_filesize.to_kb(decimal_places)
					presentation_original_filesize = presentation_original_filesize.to_kb(decimal_places)
					image_original_filesize = image_original_filesize.to_kb(decimal_places)
					drawing_original_filesize = drawing_original_filesize.to_kb(decimal_places)
					other_document_original_filesize = other_document_original_filesize.to_kb(decimal_places)
					multimedia_original_filesize = multimedia_original_filesize.to_kb(decimal_places)
					database_original_filesize = database_original_filesize.to_kb(decimal_places)
					container_original_filesize = container_original_filesize.to_kb(decimal_places)
					system_original_filesize = system_original_filesize.to_kb(decimal_places)
					no_data_original_filesize = no_data_original_filesize.to_kb(decimal_places)
					unrecognised_original_filesize = unrecognised_original_filesize.to_kb(decimal_places)
					log_original_filesize = log_original_filesize.to_kb(decimal_places)
					duplicate_filesize = duplicate_filesize.to_kb(decimal_places)
					email_duplicate_filesize = email_duplicate_filesize.to_kb(decimal_places)
					calendar_duplicate_filesize = calendar_duplicate_filesize.to_kb(decimal_places)
					contact_duplicate_filesize = contact_duplicate_filesize.to_kb(decimal_places)
					document_duplicate_filesize = document_duplicate_filesize.to_kb(decimal_places)
					spreadsheet_duplicate_filesize = spreadsheet_duplicate_filesize.to_kb(decimal_places)
					presentation_duplicate_filesize = presentation_duplicate_filesize.to_kb(decimal_places)
					image_duplicate_filesize = image_duplicate_filesize.to_kb(decimal_places)
					drawing_duplicate_filesize = drawing_duplicate_filesize.to_kb(decimal_places)
					other_document_duplicate_filesize = other_document_duplicate_filesize.to_kb(decimal_places)
					multimedia_duplicate_filesize = multimedia_duplicate_filesize.to_kb(decimal_places)
					database_duplicate_filesize = database_duplicate_filesize.to_kb(decimal_places)
					container_duplicate_filesize = container_duplicate_filesize.to_kb(decimal_places)
					system_duplicate_filesize = system_duplicate_filesize.to_kb(decimal_places)
					no_data_duplicate_filesize = no_data_duplicate_filesize.to_kb(decimal_places)
					unrecognised_duplicate_filesize = unrecognised_duplicate_filesize.to_kb(decimal_places)
					log_duplicate_filesize = log_duplicate_filesize.to_kb(decimal_places)
				when "mb"
					total_filesize = total_filesize.to_mb(decimal_places)
					email_total_filesize = email_total_filesize.to_mb(decimal_places)
					calendar_total_filesize = calendar_total_filesize.to_mb(decimal_places)
					contact_total_filesize = contact_total_filesize.to_mb(decimal_places)
					document_total_filesize = document_total_filesize.to_mb(decimal_places)
					spreadsheet_total_filesize = spreadsheet_total_filesize.to_mb(decimal_places)
					presentation_total_filesize = presentation_total_filesize.to_mb(decimal_places)
					image_total_filesize = image_total_filesize.to_mb(decimal_places)
					drawing_total_filesize = drawing_total_filesize.to_mb(decimal_places)
					other_document_total_filesize = other_document_total_filesize.to_mb(decimal_places)
					multimedia_total_filesize = multimedia_total_filesize.to_mb(decimal_places)
					database_total_filesize = database_total_filesize.to_mb(decimal_places)
					container_total_filesize = container_total_filesize.to_mb(decimal_places)
					system_total_filesize = system_total_filesize.to_mb(decimal_places)
					no_data_total_filesize = no_data_total_filesize.to_mb(decimal_places)
					unrecognised_total_filesize = unrecognised_total_filesize.to_mb(decimal_places)
					log_total_filesize = log_total_filesize.to_mb(decimal_places)
					original_filesize = original_filesize.to_mb(decimal_places)
					email_original_filesize = email_original_filesize.to_mb(decimal_places)
					calendar_original_filesize = calendar_original_filesize.to_mb(decimal_places)
					contact_original_filesize = contact_original_filesize.to_mb(decimal_places)
					document_original_filesize = document_original_filesize.to_mb(decimal_places)
					spreadsheet_original_filesize = spreadsheet_original_filesize.to_mb(decimal_places)
					presentation_original_filesize = presentation_original_filesize.to_mb(decimal_places)
					image_original_filesize = image_original_filesize.to_mb(decimal_places)
					drawing_original_filesize = drawing_original_filesize.to_mb(decimal_places)
					other_document_original_filesize = other_document_original_filesize.to_mb(decimal_places) 
					multimedia_original_filesize = multimedia_original_filesize.to_mb(decimal_places)
					database_original_filesize = database_original_filesize.to_mb(decimal_places)
					container_original_filesize = container_original_filesize.to_mb(decimal_places)
					system_original_filesize = system_original_filesize.to_mb(decimal_places)
					no_data_original_filesize = no_data_original_filesize.to_mb(decimal_places)
					unrecognised_original_filesize = unrecognised_original_filesize.to_mb(decimal_places) 
					log_original_filesize = log_original_filesize.to_mb(decimal_places)
					duplicate_filesize = duplicate_filesize.to_mb(decimal_places)
					email_duplicate_filesize = email_duplicate_filesize.to_mb(decimal_places)
					calendar_duplicate_filesize = calendar_duplicate_filesize.to_mb(decimal_places)
					contact_duplicate_filesize = contact_duplicate_filesize.to_mb(decimal_places)
					document_duplicate_filesize = document_duplicate_filesize.to_mb(decimal_places)
					spreadsheet_duplicate_filesize = spreadsheet_duplicate_filesize.to_mb(decimal_places)
					presentation_duplicate_filesize = presentation_duplicate_filesize.to_mb(decimal_places)
					image_duplicate_filesize = image_duplicate_filesize.to_mb(decimal_places)
					drawing_duplicate_filesize = drawing_duplicate_filesize.to_mb(decimal_places)
					other_document_duplicate_filesize = other_document_duplicate_filesize.to_mb(decimal_places)
					multimedia_duplicate_filesize = multimedia_duplicate_filesize.to_mb(decimal_places)
					database_duplicate_filesize = database_duplicate_filesize.to_mb(decimal_places)
					container_duplicate_filesize = container_duplicate_filesize.to_mb(decimal_places)
					system_duplicate_filesize = system_duplicate_filesize.to_mb(decimal_places)
					no_data_duplicate_filesize = no_data_duplicate_filesize.to_mb(decimal_places)
					unrecognised_duplicate_filesize = unrecognised_duplicate_filesize.to_mb(decimal_places)
					log_duplicate_filesize = log_duplicate_filesize.to_mb(decimal_places)
				when "gb"
					total_filesize = total_filesize.to_gb(decimal_places)
					email_total_filesize = email_total_filesize.to_gb(decimal_places)
					calendar_total_filesize = calendar_total_filesize.to_gb(decimal_places)
					contact_total_filesize = contact_total_filesize.to_gb(decimal_places)
					document_total_filesize = document_total_filesize.to_gb(decimal_places)
					spreadsheet_total_filesize = spreadsheet_total_filesize.to_gb(decimal_places)
					presentation_total_filesize = presentation_total_filesize.to_gb(decimal_places)
					image_total_filesize = image_total_filesize.to_gb(decimal_places)
					drawing_total_filesize = drawing_total_filesize.to_gb(decimal_places)
					other_document_total_filesize = other_document_total_filesize.to_gb(decimal_places)
					multimedia_total_filesize = multimedia_total_filesize.to_gb(decimal_places)
					database_total_filesize = database_total_filesize.to_gb(decimal_places)
					container_total_filesize = container_total_filesize.to_gb(decimal_places)
					system_total_filesize = system_total_filesize.to_gb(decimal_places)
					no_data_total_filesize = no_data_total_filesize.to_gb(decimal_places)
					unrecognised_total_filesize = unrecognised_total_filesize.to_gb(decimal_places)
					log_total_filesize = log_total_filesize.to_gb(decimal_places)
					original_filesize = original_filesize.to_gb(decimal_places)
					email_original_filesize = email_original_filesize.to_gb(decimal_places)
					calendar_original_filesize = calendar_original_filesize.to_gb(decimal_places)
					contact_original_filesize = contact_original_filesize.to_gb(decimal_places)
					document_original_filesize = document_original_filesize.to_gb(decimal_places)
					spreadsheet_original_filesize = spreadsheet_original_filesize.to_gb(decimal_places)
					presentation_original_filesize = presentation_original_filesize.to_gb(decimal_places)
					image_original_filesize = image_original_filesize.to_gb(decimal_places)
					drawing_original_filesize = drawing_original_filesize.to_gb(decimal_places)
					other_document_original_filesize = other_document_original_filesize.to_gb(decimal_places)
					multimedia_original_filesize = multimedia_original_filesize.to_gb(decimal_places)
					database_original_filesize = database_original_filesize.to_gb(decimal_places)
					container_original_filesize = container_original_filesize.to_gb(decimal_places)
					system_original_filesize = system_original_filesize.to_gb(decimal_places)
					no_data_original_filesize = no_data_original_filesize.to_gb(decimal_places)
					unrecognised_original_filesize = unrecognised_original_filesize.to_gb(decimal_places)
					log_original_filesize = log_original_filesize.to_gb(decimal_places)
					duplicate_filesize = duplicate_filesize.to_gb(decimal_places)
					email_duplicate_filesize = email_duplicate_filesize.to_gb(decimal_places)
					calendar_duplicate_filesize = calendar_duplicate_filesize.to_gb(decimal_places)
					contact_duplicate_filesize = contact_duplicate_filesize.to_gb(decimal_places)
					document_duplicate_filesize = document_duplicate_filesize.to_gb(decimal_places)
					spreadsheet_duplicate_filesize = spreadsheet_duplicate_filesize.to_gb(decimal_places)
					presentation_duplicate_filesize = presentation_duplicate_filesize.to_gb(decimal_places)
					image_duplicate_filesize = image_duplicate_filesize.to_gb(decimal_places)
					drawing_duplicate_filesize = drawing_duplicate_filesize.to_gb(decimal_places)
					other_document_duplicate_filesize = other_document_duplicate_filesize.to_gb(decimal_places)
					multimedia_duplicate_filesize = multimedia_duplicate_filesize.to_gb(decimal_places)
					database_duplicate_filesize = database_duplicate_filesize.to_gb(decimal_places)
					container_duplicate_filesize = container_duplicate_filesize.to_gb(decimal_places)
					system_duplicate_filesize = system_duplicate_filesize.to_gb(decimal_places)
					no_data_duplicate_filesize = no_data_duplicate_filesize.to_gb(decimal_places)
					unrecognised_duplicate_filesize = unrecognised_duplicate_filesize.to_gb(decimal_places)
					log_duplicate_filesize = log_duplicate_filesize.to_gb(decimal_places)
				else
					total_filesize = total_filesize
					email_total_filesize = email_total_filesize  
					calendar_total_filesize = calendar_total_filesize 
					contact_total_filesize = contact_total_filesize
					document_total_filesize = document_total_filesize
					spreadsheet_total_filesize = spreadsheet_total_filesize 
					presentation_total_filesize = presentation_total_filesize  
					image_total_filesize = image_total_filesize 
					drawing_total_filesize = drawing_total_filesize 
					other_document_total_filesize = other_document_total_filesize 
					multimedia_total_filesize = multimedia_total_filesize 
					database_total_filesize = database_total_filesize 
					container_total_filesize = container_total_filesize 
					system_total_filesize = system_total_filesize 
					no_data_total_filesize = no_data_total_filesize 
					unrecognised_total_filesize = unrecognised_total_filesize 
					log_total_filesize = log_total_filesize 
					original_filesize = original_filesize 
					email_original_filesize = email_original_filesize 
					calendar_original_filesize = calendar_original_filesize 
					contact_original_filesize = contact_original_filesize 
					document_original_filesize = document_original_filesize 
					spreadsheet_original_filesize = spreadsheet_original_filesize 
					presentation_original_filesize = presentation_original_filesize 
					image_original_filesize = image_original_filesize 
					drawing_original_filesize = drawing_original_filesize 
					other_document_original_filesize = other_document_original_filesize 
					multimedia_original_filesize = multimedia_original_filesize 
					database_original_filesize = database_original_filesize 
					container_original_filesize = container_original_filesize 
					system_original_filesize = system_original_filesize 
					no_data_original_filesize = no_data_original_filesize 
					unrecognised_original_filesize = unrecognised_original_filesize 
					log_original_filesize = log_original_filesize 
					duplicate_filesize = duplicate_filesize 
					email_duplicate_filesize = email_duplicate_filesize 
					calendar_duplicate_filesize = calendar_duplicate_filesize 
					contact_duplicate_filesize = contact_duplicate_filesize
					document_duplicate_filesize = document_duplicate_filesize 
					spreadsheet_duplicate_filesize = spreadsheet_duplicate_filesize 
					presentation_duplicate_filesize = presentation_duplicate_filesize 
					image_duplicate_filesize = image_duplicate_filesize 
					drawing_duplicate_filesize = drawing_duplicate_filesize 
					other_document_duplicate_filesize = other_document_duplicate_filesize 
					multimedia_duplicate_filesize = multimedia_duplicate_filesize 
					database_duplicate_filesize = database_duplicate_filesize 
					container_duplicate_filesize = container_duplicate_filesize 
					system_duplicate_filesize = system_duplicate_filesize 
					no_data_duplicate_filesize = no_data_duplicate_filesize 
					unrecognised_duplicate_filesize = unrecognised_duplicate_filesize
					log_duplicate_filesize = log_duplicate_filesize 
			end
			# store all sizes 
			totals_count = 'Total::' + "#{total_count}" + "::#{total_filesize}" + ';Email::' + "#{email_total_count}" + "::#{email_total_filesize}" + ';Calendar::' + "#{calendar_total_count}" + "::#{calendar_total_filesize}" + ';Contact::' + "#{contact_total_count}" + "::#{contact_total_filesize}" + ';Document::' + "#{document_total_count}" + "::#{document_total_filesize}" + ';Spreadsheet::' + "#{spreadsheet_total_count}" + "::#{spreadsheet_total_filesize}" + ';Presentation::' + "#{presentation_total_count}" + "::#{presentation_total_filesize}" + ';Image::' + "#{image_total_count}" + "::#{image_total_filesize}" + ';Drawing::' + "#{drawing_total_count}" + "::#{drawing_total_filesize}" + ';Other-Document::' + "#{other_document_total_count}" + "::#{other_document_total_filesize}" + ';Multimedia::' + "#{multimedia_total_count}" + "::#{multimedia_total_filesize}" + ';Database::' + "#{database_total_count}" + "::#{database_total_filesize}" + ';Container::' + "#{container_total_count}" + "::#{container_total_filesize}" + ';System::' + "#{system_total_count}" + "::#{system_total_filesize}" + ';No-data::' + "#{no_data_total_count}" + "::#{no_data_total_filesize}" + ';Unrecognised::' + "#{unrecognised_total_count}" + "::#{unrecognised_total_filesize}" + ';Log::' + "#{log_total_count}" + "::#{log_total_filesize}" + ';Chat-Conversations::' + "#{chat_conversation_total_count}" + "::#{chat_conversation_total_filesize}" + ';Chat-Messages::' + "#{chat_message_total_count}" + "::#{chat_message_total_filesize}"
			originals_count = 'Total::' + "#{original_count}" + "::#{original_filesize}" + ';Email::' + "#{email_original_count}" + "::#{email_original_filesize}" + ';Calendar::' + "#{calendar_original_count}" + "::#{calendar_original_filesize}" + ';Contact::' + "#{contact_original_count}" + "::#{contact_original_filesize}" + ';Document::' + "#{document_original_count}" + "::#{document_original_filesize}" + ';Spreadsheet::' + "#{spreadsheet_original_count}" + "::#{spreadsheet_original_filesize}" + ';Presentation::' + "#{presentation_original_count}" + "::#{presentation_original_filesize}" + ';Image::' + "#{image_original_count}" + "::#{image_original_filesize}" + ';Drawing::' + "#{drawing_original_count}" + "::#{drawing_original_filesize}" + ';Other-Document::' + "#{other_document_original_count}" + "::#{other_document_original_filesize}" + ';Multimedia::' + "#{multimedia_original_count}" + "::#{multimedia_original_filesize}" + ';Database::' + "#{database_original_count}" + "::#{database_original_filesize}" + ';Container::' + "#{container_original_count}" + "::#{container_original_filesize}" + ';System::' + "#{system_original_count}" + "::#{system_original_filesize}" + ';No-data::' + "#{no_data_original_count}" + "::#{no_data_original_filesize}" + ';Unrecognised::' + "#{unrecognised_original_count}" + "::#{unrecognised_original_filesize}" + ';Log::' + "#{log_original_count}" + "::#{log_original_filesize}" + ';Chat-Conversations::' + "#{chat_conversation_original_count}" + "::#{chat_conversation_original_filesize}"  + ';Chat-Messages::' + "#{chat_message_original_count}" + "::#{chat_message_original_filesize}"
			duplicates_count = 'Total::' + "#{duplicate_count}" + "::#{duplicate_filesize}" + ';Email::' + "#{email_duplicate_count}" + "::#{email_duplicate_filesize}" + ';Calendar::' + "#{calendar_duplicate_count}" + "::#{calendar_duplicate_filesize}" + ';Contact::' + "#{contact_duplicate_count}" + "::#{contact_duplicate_filesize}" + ';Document::' + "#{document_duplicate_count}" + "::#{document_duplicate_filesize}" + ';Spreadsheet::' + "#{spreadsheet_duplicate_count}" + "::#{spreadsheet_duplicate_filesize}" + ';Presentation::' + "#{presentation_duplicate_count}" + "::#{presentation_duplicate_filesize}" + ';Image::' + "#{image_duplicate_count}" + "::#{image_duplicate_filesize}" + ';Drawing::' + "#{drawing_duplicate_count}" + "::#{drawing_duplicate_filesize}" + ';Other-Document::' + "#{other_document_duplicate_count}" + "::#{other_document_duplicate_filesize}" + ';Multimedia::' + "#{multimedia_duplicate_count}" + "::#{multimedia_duplicate_filesize}" + ';Database::' + "#{database_duplicate_count}" + "::#{database_duplicate_filesize}" + ';Container::' + "#{container_duplicate_count}" + "::#{container_duplicate_filesize}" + ';System::' + "#{system_duplicate_count}" + "::#{system_duplicate_filesize}" + ';No-data::' + "#{no_data_duplicate_count}" + "::#{no_data_duplicate_filesize}" + ';Unrecognised::' + "#{unrecognised_duplicate_count}" + "::#{unrecognised_duplicate_filesize}" + ';Log::' + "#{log_duplicate_count}" + "::#{log_duplicate_filesize}" + ';Chat-Conversations::' + "#{chat_conversation_duplicate_count}" + "::#{chat_conversation_duplicate_filesize}"  + ';Chat-Messages::' + "#{chat_message_duplicate_count}" + "::#{chat_message_duplicate_filesize}"

			puts "Totals - #{totals_count}"
			puts "Originals - #{originals_count}"
			puts "Duplicates - #{duplicates_count}"

			# update the database with the data collected to this point
			case_items_data = [totals_count, totals_count, originals_count, duplicates_count, 35]
			db.update("UPDATE NuixReportingInfo SET ItemTypes = ?, ItemCounts = ?, OriginalItems = ?, DuplicateItems = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", case_items_data)
		rescue StandardError => msg
			puts "Error in getCaseSize + #{msg}"
		ensure

		end

	end
	
	# get all the types from the case
	def self.getCaseTypes(nuix_case, case_guid, showsizein, decimal_places, db)
		begin
			#get items types
			item_types = nuix_case.getItemTypes
			mime_type_count = ''
			# loop through each item type and get the sizes
			item_types.each do |item_search|
				mime_type_size = nuix_case.getStatistics.getFileSize("mime-type:#{item_search}")
				case showsizein
					when "byte"
						mime_type_size = mime_type_size
					when "kb"
						mime_type_size = mime_type_size.to_kb(decimal_places)
					when "mb"
						mime_type_size = mime_type_size.to_mb(decimal_places)
					when "gb"
						mime_type_size = mime_type_size.to_gb(decimal_places)
					else
						mime_type_size = mime_type_size
				end

				mime_type_count = mime_type_count + "#{item_search}::" + nuix_case.count("mime-type:#{item_search}").to_s + "::" + mime_type_size.to_s + ";"
			end
			#update the database with the information collected to this point
			case_mimetype_data = [mime_type_count, 40]
			db.update("UPDATE NuixReportingInfo SET MimeTypes = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", case_mimetype_data)
		rescue StandardError => msg
			puts "Error in getCaseTypes + #{msg}"
		ensure

		end
	end
	
	# get the case history from the case
	def self.getCaseHistory(nuix_case, case_guid, showsizein, decimal_places, includeannotations, db)
		begin
			date_range_start = nil
			date_range_end = nil
			user = ''
			open_sessions = []
			close_sessions = []
			loadData_sessions = []
			open_sorted = []
			close_sorted = []
			sorted_array = []
			session_data = []
			user_sessions = []
			session_users = []
			loaddata_users = []
			time_diff = 0
			start_datetime = ''
			end_datetime = ''
			loadEvents = 0
			exportEvents = 0
			scriptEvents = 0
			scriptdata_users = ''
			exportdata_users = ''
			total_loadTime = 0
			db.update("Delete from UCRTSessionEvents where CaseGUID = '#{case_guid}'")

			open_options = {
				"startDateAfter"=> date_range_start,
				"startDateBefore"=> date_range_end,
				"type"=> "openSession",
			}

			close_options = {
				"startDateAfter"=> date_range_start,
				"startDateBefore"=> date_range_end,
				"type"=> "closeSession",
			}
			load_options = {
				"startDateAfter" => date_range_start,
				"startDateBefore"=> date_range_end,
				"type" => "loadData",
			}
			script_options = {
				"startDateAfter" => date_range_start,
				"startDateBefore"=> date_range_end,
				"type" => "script",
			}
			export_options = {
				"startDateAfter" => date_range_start,
				"startDateBefore"=> date_range_end,
				"type" => "export",
			}
			annotate_options = {
				"startDateAfter" => date_range_start,
				"startDateBefore"=> date_range_end,
				"type" => "annotation",
			}
			#Build history options hash
			options = {
				"startDateAfter"=> date_range_start,
				"startDateBefore"=> date_range_end,
				"user"=> user,
			}
			openhistory = nuix_case.getHistory(open_options)
			openhistory.each do |openhist|
				username = openhist.getUser
				open_datetime = openhist.getStartDate
				end_datetime = openhist.getEndDate
				event = openhist.getTypeString
				session_data = [case_guid, event, open_datetime, end_datetime, '', username]
				db.update("INSERT INTO UCRTSessionEvents (CaseGUID, SessionEvent , StartDate , EndDate, Duration, User ) VALUES ( ?, ?, ?, ?, ?, ? )", session_data)
			end

			closehistory = nuix_case.getHistory(close_options)
			closehistory.each do |closehist|
				username = closehist.getUser
				open_datetime = closehist.getStartDate
				end_datetime = closehist.getEndDate
				event = closehist.getTypeString
				session_data = [case_guid, event, open_datetime, end_datetime, '', username]
				db.update("INSERT INTO UCRTSessionEvents (CaseGUID, SessionEvent , StartDate , EndDate, Duration, User ) VALUES ( ?, ?, ?, ?, ?, ? )", session_data)
			end

			loadDataHistory = nuix_case.getHistory(load_options)
			loadDataHistory.each do |loadhist|
				username = loadhist.getUser
				start_datetime = loadhist.getStartDate
				end_datetime = loadhist.getEndDate
				total_time = end_datetime.getMillis - start_datetime.getMillis
				total_time = total_time / 1000
				event = loadhist.getTypeString
				total_loadTime = total_loadTime + total_time
				loadEvents = loadEvents + 1
				total_time = TimeSpanFormatter.format_seconds(total_time)
				session_data = [case_guid, event, start_datetime, end_datetime, total_time, username]
				db.update("INSERT INTO UCRTSessionEvents (CaseGUID, SessionEvent , StartDate , EndDate, Duration, User ) VALUES ( ?, ?, ?, ?, ?, ? )", session_data)
			end
			puts "Load Data Start = #{start_datetime}"
			puts "Load Data End = #{end_datetime}"
			total_time = TimeSpanFormatter.format_seconds(total_loadTime)
			case_loadtime_data = [total_loadTime, start_datetime, end_datetime, loadEvents, 50]
			db.update("UPDATE NuixReportingInfo SET TotalLoadTime = ?, LoadDataStart = ?, LoadDataEnd = ?, LoadEvents = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", case_loadtime_data)

			scriptRunHistory = nuix_case.getHistory(script_options)
			scriptRunHistory.each do |scripthist|
				username = scripthist.getUser
				open_datetime = scripthist.getStartDate
				end_datetime = scripthist.getEndDate
				event = scripthist.getTypeString
				session_data = [case_guid, event, open_datetime, end_datetime, '', username]
				db.update("INSERT INTO UCRTSessionEvents (CaseGUID, SessionEvent , StartDate , EndDate, Duration, User ) VALUES ( ?, ?, ?, ?, ?, ? )", session_data)
			end

			if includeannotations == true
				annotateHistory = nuix_case.getHistory(annotate_options)
				annotateHistory.each do |annotatehist|
					username = annotatehist.getUser
					open_datetime = annotatehist.getStartDate
					end_datetime = annotatehist.getEndDate
					event = annotatehist.getTypeString
					session_data = [case_guid, event, open_datetime, end_datetime, '', username]
					db.update("INSERT INTO UCRTSessionEvents (CaseGUID, SessionEvent , StartDate , EndDate, Duration, User ) VALUES ( ?, ?, ?, ?, ?, ? )", session_data)
				end
			end 
			exportHistory = nuix_case.getHistory(export_options)
			exportHistory.each do |exporthist|
				exportdetails = exporthist.getDetails
				success_items = exportdetails["processed"]
				failed_items = exportdetails["failed"]
				username = exporthist.getUser
				open_datetime = exporthist.getStartDate
				end_datetime = exporthist.getEndDate
				event = exporthist.getTypeString
				session_data = [case_guid, event, open_datetime, end_datetime, '', username, success_items, failed_items]
				db.update("INSERT INTO UCRTSessionEvents (CaseGUID, SessionEvent , StartDate , EndDate, Duration, User, Success, Failures ) VALUES ( ?, ?, ?, ?, ?, ?, ?, ? )", session_data)
			end

		rescue StandardError => msg
			puts "Error in getCaseHistory + #{msg}"
		ensure

		end
	end
	# get the date range information if necessary - this method will iterate through each day from
	# the earliest item in the case to the latest item in the case and get case counts for each day
	# for example earliest item 1-1-2022 - latest item 12-31-2022 will loop 365 times for each item type for each day of the year
	def self.getDateRangeInfo(nuix_case, case_guid, db)
		begin
			counter = 0
			kinds_count = ''
			email_type_size = 0
			email_zantaz_type_size = 0
			calendar_type_size = 0
			contact_type_size = 0
			documents_type_size = 0
			spreadsheets_type_size = 0
			presentations_type_size = 0
			image_type_size = 0
			drawings_type_size = 0
			otherdocuments_type_size = 0
			multimedia_type_size = 0
			database_type_size = 0
			container_type_size = 0
			system_type_size = 0
			nodata_type_size = 0
			unrecognised_type_size = 0
			log_type_size = 0

			email_total_count = 0
			email_total_size = 0
			calendar_total_count = 0
			calendar_total_size = 0
			contact_total_count = 0
			contact_total_size = 0
			documents_total_count = 0
			documents_total_size = 0
			spreadsheets_total_count = 0
			spreadsheets_total_size = 0
			presentations_total_count = 0
			presentations_total_size = 0
			image_total_count = 0
			image_total_size = 0
			drawings_total_count = 0
			drawings_total_size = 0
			otherdocuments_total_count = 0
			otherdocuments_total_size = 0
			multimedia_total_count = 0
			multimedia_total_size = 0
			database_total_count = 0
			database_total_size = 0
			container_total_count = 0
			container_total_size = 0
			system_total_count = 0
			system_total_size = 0
			nodata_total_count = 0
			nodata_total_size = 0
			unrecognised_total_count = 0
			unrecognised_total_size = 0
			log_total_count = 0
			log_total_size = 0

			daterange = nuix_case.getStatistics.getCaseDateRange
			oldest = daterange.getEarliest
			newest = daterange.getLatest
			oldest={"defaultFields"=>"item-date","order"=>"item-date ASC","limit"=>1}
			newest={"defaultFields"=>"item-date","order"=>"item-date DESC","limit"=>1}
			oldestItemDate = nuix_case.search("flag:top_level item-date:*", oldest).first().getDate().to_s
			newestItemDate = nuix_case.search("flag:top_level item-date:*", newest).first().getDate().to_s
			puts "Oldest Item date = #{oldestItemDate}"
			puts "Newest Item date = #{newestItemDate}"

			currentdate = Date.today
			currentdate = currentdate.strftime("%Y%m%d")

			if newestItemDate != ''
				newestDate = Date.parse(newestItemDate.to_s.split("T").first)
				if newestDate > Date.today
					search_bad_date = "item-date:[#{currentdate} TO *]"
					baddateitemcount = nuix_case.count(search_bad_date)
					insert_data = [case_guid, "items", "1/1/3000", baddateitemcount, "NA"]
	#				db.update("Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)", insert_data)
					newestDate = Date.today
				end
			end
			oldestDate= Date.parse(oldestItemDate.to_s.split("T").first)
			commit_size = 2500
			db.batch_insert("Insert into UCRTDateRange(CaseGUID,ItemType,ItemDate,ItemCount, Custodian) VALUES (?, ?, ?, ?, ?)", commit_size) do |batch|
				oldestDate.upto(newestDate) do |dateitem|
					searchstart = DateTime.now
					searchstarttime = searchstart.strftime("%d/%m/%Y %H:%M")

					searchdate = dateitem.to_s
					dateitemsearch = searchdate.to_s.delete('-')
					anyitem_search = "item-date:#{dateitemsearch}"
					anyitem_count = nuix_case.count(anyitem_search)
					if anyitem_count > 0
						puts "Searching data for date #{dateitemsearch} - has #{anyitem_count} items"
						email_search_term = "item-date:#{dateitemsearch} and kind:email and flag:top_level and has-custodian:0 and not mime-type:application/vnd.zantaz.archive"
						email_type_count = nuix_case.count(email_search_term)
						if email_type_count > 0
							email_total_size += nuix_case.getStatistics.getFileSize(email_search_term)
						end
						insert_data = [case_guid, "email", searchdate, email_type_count, "NA"]
						batch.insert(insert_data)

						email_zantaz_search_term = "item-date:#{dateitemsearch} and kind:email and flag:top_level and has-custodian:0 and mime-type:application/vnd.zantaz.archive"
						email_zantaz_type_count = nuix_case.count(email_zantaz_search_term)
						if email_zantaz_type_count > 0
							email_total_size += nuix_case.getStatistics.getFileSize(email_search_term)
						end
						insert_data = [case_guid, "email-zantaz", searchdate, email_zantaz_type_count, "NA"]
						batch.insert(insert_data)

						calendar_search_term = "item-date:#{dateitemsearch} and kind:calendar and has-custodian:0"
						calendar_type_count = nuix_case.count(calendar_search_term)
						if calendar_type_count > 0
							calendar_total_size += nuix_case.getStatistics.getFileSize(calendar_search_term)
						end
						insert_data = [case_guid, "calendar", searchdate, calendar_type_count, "NA"]
						batch.insert(insert_data)

						contact_search_term = "item-date:#{dateitemsearch} and kind:contact and has-custodian:0"
						contact_type_count = nuix_case.count(contact_search_term)
						if contact_type_count > 0
							contact_total_size += nuix_case.getStatistics.getFileSize(contact_search_term)
						end
						insert_data = [case_guid, "contact", searchdate, contact_type_count, "NA"]
						batch.insert(insert_data)
						
						document_search_term = "item-date:#{dateitemsearch} and kind:document and has-custodian:0"
						documents_type_count = nuix_case.count(document_search_term)
						if documents_type_count > 0
							documents_total_size += nuix_case.getStatistics.getFileSize(document_search_term)
						end
						insert_data = [case_guid, "document", searchdate, documents_type_count, "NA"]
						batch.insert(insert_data)

						spreadsheet_search_term = "item-date:#{dateitemsearch} and kind:spreadsheet and has-custodian:0"
						spreadsheets_type_count = nuix_case.count(spreadsheet_search_term)
						if spreadsheets_type_count > 0
							spreadsheets_total_size += nuix_case.getStatistics.getFileSize(spreadsheet_search_term)
						end
						insert_data = [case_guid, "spreadsheet", searchdate, spreadsheets_type_count, "NA"]
						batch.insert(insert_data)

						presentation_search_term = "item-date:#{dateitemsearch} and kind:presentation and has-custodian:0"
						presentations_type_count = nuix_case.count(presentation_search_term)
						if presentations_type_count > 0
							presentations_total_size += nuix_case.getStatistics.getFileSize(presentation_search_term)
						end
						insert_data = [case_guid, "presentation", searchdate, presentations_type_count, "NA"]
						batch.insert(insert_data)

						image_search_term = "item-date:#{dateitemsearch} and kind:image and has-custodian:0"
						image_type_count = nuix_case.count(image_search_term)
						if image_type_count > 0
							image_total_size += nuix_case.getStatistics.getFileSize(image_search_term)
						end
						insert_data = [case_guid, "image", searchdate, image_type_count, "NA"]
						batch.insert(insert_data)

						drawing_search_term = "item-date:#{dateitemsearch} and kind:drawing and has-custodian:0"
						drawings_type_count = nuix_case.count(drawing_search_term)
						if drawings_type_count > 0
							drawings_total_size += nuix_case.getStatistics.getFileSize(drawing_search_term)
						end
						insert_data = [case_guid, "drawing", searchdate, drawings_type_count, "NA"]
						batch.insert(insert_data)

						otherdocument_search_term = "item-date:#{dateitemsearch} and kind:other-document and has-custodian:0"
						otherdocuments_type_count = nuix_case.count(otherdocument_search_term)
						if otherdocuments_type_count > 0
							otherdocuments_total_size += nuix_case.getStatistics.getFileSize(otherdocument_search_term)
						end
						insert_data = [case_guid, "otherdocument", searchdate, otherdocuments_type_count, "NA"]
						batch.insert(insert_data)

						multimedia_search_term = "item-date:#{dateitemsearch} and kind:multimedia and has-custodian:0"
						multimedia_type_count = nuix_case.count(multimedia_search_term)
						if multimedia_type_count > 0
							multimedia_total_size += nuix_case.getStatistics.getFileSize(multimedia_search_term)
						end
						insert_data = [case_guid, "multimedia", searchdate, multimedia_type_count, "NA"]
						batch.insert(insert_data)

						database_search_term = "item-date:#{dateitemsearch} and kind:database and has-custodian:0"
						database_type_count = nuix_case.count(database_search_term)
						if database_type_count > 0
							database_total_size += nuix_case.getStatistics.getFileSize(database_search_term)
						end
						insert_data = [case_guid, "database", searchdate, database_type_count, "NA"]
						batch.insert(insert_data)

						container_search_term = "item-date:#{dateitemsearch} and kind:container AND NOT mime-type:application/vnd.nuix-evidence and has-custodian:0"
						container_type_count = nuix_case.count(container_search_term)
						if container_type_count > 0
							container_total_size += nuix_case.getStatistics.getFileSize(container_search_term)
						end
						insert_data = [case_guid, "container", searchdate, container_type_count, "NA"]
						batch.insert(insert_data)

						system_search_term = "item-date:#{dateitemsearch} and kind:system and has-custodian:0"
						system_type_count = nuix_case.count(system_search_term)
						if system_type_count > 0
							system_total_size += nuix_case.getStatistics.getFileSize(system_search_term)
						end
						insert_data = [case_guid, "system", searchdate, system_type_count, "NA"]
						batch.insert(insert_data)

						nodata_search_term = "item-date:#{dateitemsearch} and kind:no-data and has-custodian:0"
						nodata_type_count = nuix_case.count(nodata_search_term)
						if nodata_type_count > 0
							nodata_total_size += nuix_case.getStatistics.getFileSize(nodata_search_term)
						end
						insert_data = [case_guid, "nodata", searchdate, nodata_type_count, "NA"]
						batch.insert(insert_data)

						unrecognised_search_term = "item-date:#{dateitemsearch} and kind:unrecognised and has-custodian:0"
						unrecognised_type_count = nuix_case.count(unrecognised_search_term)
						if unrecognised_type_count > 0
							unrecognised_total_size += nuix_case.getStatistics.getFileSize(unrecognised_search_term)
						end
						insert_data = [case_guid, "unrecognised", searchdate, unrecognised_type_count, "NA"]
						batch.insert(insert_data)

						log_search_term = "item-date:#{dateitemsearch} and kind:log and has-custodian:0"
						log_type_count = nuix_case.count(log_search_term)
						if log_type_count > 0
							log_total_size += nuix_case.getStatistics.getFileSize(log_search_term)
						end
						insert_data = [case_guid, "log", searchdate, log_type_count, "NA"]
						batch.insert(insert_data)

						custodian_count = 0
						all_custodians_info = ''
						custodian_names = nuix_case.getAllCustodians
						custodian_names.each do |custodian_name|
							anyitem_custodian_search = "item-date:#{dateitemsearch} and custodian:'#{custodian_name}'"
							anyitem_custodian_count = nuix_case.count(anyitem_custodian_search)
							if anyitem_custodian_count > 0
								email_search_term = "item-date:#{dateitemsearch} and kind:email and flag:top_level and custodian:'#{custodian_name}' and not mime-type:application/vnd.zantaz.archive"
								email_type_count_cust = nuix_case.count(email_search_term)
								if email_type_count_cust > 0
									email_total_size += nuix_case.getStatistics.getFileSize(email_search_term)
								end
								email_total_count += email_type_count_cust
								insert_data = [case_guid, "email", searchdate, email_type_count_cust, custodian_name]
								batch.insert(insert_data)

								email_zantaz_search_term = "item-date:#{dateitemsearch} and kind:email and flag:top_level and custodian:'#{custodian_name}' and mime-type:application/vnd.zantaz.archive"
								email_zantaz_type_count_cust = nuix_case.count(email_zantaz_search_term)
								if email_zantaz_type_count_cust > 0
									email_total_size += nuix_case.getStatistics.getFileSize(email_zantaz_search_term)
								end
								email_total_count += email_zantaz_type_count_cust
								insert_data = [case_guid, "email_zantaz", searchdate, email_zantaz_type_count_cust, custodian_name]
								batch.insert(insert_data)

								calendar_search_term = "item-date:#{dateitemsearch} and kind:calendar and custodian:'#{custodian_name}'"
								calendar_type_count_cust = nuix_case.count(calendar_search_term)
								if calendar_type_count_cust > 0
									calendar_total_size += nuix_case.getStatistics.getFileSize(calendar_search_term)
								end
								calendar_total_count += calendar_type_count_cust
								insert_data = [case_guid, "calendar", searchdate, calendar_type_count_cust, custodian_name]
								batch.insert(insert_data)

								contact_search_term = "item-date:#{dateitemsearch} and kind:contact and custodian:'#{custodian_name}'"
								contact_type_count_cust = nuix_case.count(contact_search_term)
								if contact_type_count_cust > 0
									contact_total_size += nuix_case.getStatistics.getFileSize(contact_search_term)
								end
								contact_total_count += contact_type_count_cust
								insert_data = [case_guid, "contact", searchdate, contact_type_count_cust, custodian_name]
								batch.insert(insert_data)

								document_search_term = "item-date:#{dateitemsearch} and kind:document and custodian:'#{custodian_name}'"
								documents_type_count_cust = nuix_case.count(document_search_term)
								if documents_type_count_cust > 0
									documents_total_size += nuix_case.getStatistics.getFileSize(document_search_term)
								end
								documents_total_count += documents_type_count_cust
								insert_data = [case_guid, "document", searchdate, documents_type_count_cust, custodian_name]
								batch.insert(insert_data)

								spreadsheet_search_term = "item-date:#{dateitemsearch} and kind:spreadsheet and custodian:'#{custodian_name}'"
								spreadsheets_type_count_cust = nuix_case.count(spreadsheet_search_term)
								if spreadsheets_type_count_cust > 0
									spreadsheets_total_size += nuix_case.getStatistics.getFileSize(spreadsheet_search_term)
								end
								spreadsheets_total_count += spreadsheets_type_count_cust
								insert_data = [case_guid, "spreadsheet", searchdate, spreadsheets_type_count_cust, custodian_name]
								batch.insert(insert_data)

								presentation_search_term = "item-date:#{dateitemsearch} and kind:presentation and custodian:'#{custodian_name}'"
								presentations_type_count_cust = nuix_case.count(presentation_search_term)
								if presentations_type_count_cust > 0
									presentations_total_size += nuix_case.getStatistics.getFileSize(presentation_search_term)
								end
								presentations_total_count += presentations_type_count_cust
								insert_data = [case_guid, "presentation", searchdate, presentations_type_count_cust, custodian_name]
								batch.insert(insert_data)

								image_search_term = "item-date:#{dateitemsearch} and kind:image and custodian:'#{custodian_name}'"
								image_type_count_cust = nuix_case.count(image_search_term)
								if image_type_count_cust > 0
									image_total_size += nuix_case.getStatistics.getFileSize(image_search_term)
								end
								image_total_count += image_type_count_cust
								insert_data = [case_guid, "image", searchdate, image_type_count_cust, custodian_name]
								batch.insert(insert_data)

								drawing_search_term = "item-date:#{dateitemsearch} and kind:drawing and custodian:'#{custodian_name}'"
								drawings_type_count_cust = nuix_case.count(drawing_search_term)
								if drawings_type_count_cust > 0
									drawings_total_size += nuix_case.getStatistics.getFileSize(drawing_search_term)
								end
								drawings_total_count += drawings_type_count_cust 
								insert_data = [case_guid, "drawing", searchdate, drawings_type_count_cust, custodian_name]
								batch.insert(insert_data)

								otherdocument_search_term = "item-date:#{dateitemsearch} and kind:other-document and custodian:'#{custodian_name}'"
								otherdocuments_type_count_cust = nuix_case.count(otherdocument_search_term)
								if otherdocuments_type_count_cust > 0
									otherdocuments_total_size += nuix_case.getStatistics.getFileSize(otherdocument_search_term)
								end
								otherdocuments_total_count += otherdocuments_type_count_cust 
								insert_data = [case_guid, "otherdocument", searchdate, otherdocuments_type_count_cust, custodian_name]
								batch.insert(insert_data)

								multimedia_search_term = "item-date:#{dateitemsearch} and kind:multimedia and custodian:'#{custodian_name}'"
								multimedia_type_count_cust = nuix_case.count(multimedia_search_term)
								if multimedia_type_count_cust > 0
									multimedia_total_size += nuix_case.getStatistics.getFileSize(multimedia_search_term)
								end
								multimedia_total_count += multimedia_type_count_cust 	
								insert_data = [case_guid, "multimedia", searchdate, multimedia_type_count_cust, custodian_name]
								batch.insert(insert_data)

								database_search_term = "item-date:#{dateitemsearch} and kind:database and custodian:'#{custodian_name}'"
								database_type_count_cust = nuix_case.count(database_search_term)
								if database_type_count_cust > 0
									database_total_size += nuix_case.getStatistics.getFileSize(database_search_term)
								end
								database_total_count += database_type_count_cust
								insert_data = [case_guid, "database", searchdate, database_type_count_cust, custodian_name]
								batch.insert(insert_data)

								container_search_term = "item-date:#{dateitemsearch} and kind:container AND NOT mime-type:application/vnd.nuix-evidence and custodian:'#{custodian_name}'"
								container_type_count_cust = nuix_case.count(container_search_term)
								if container_type_count_cust > 0
									container_total_size += nuix_case.getStatistics.getFileSize(container_search_term)
								end
								container_total_count += contact_type_count_cust
								insert_data = [case_guid, "container", searchdate, container_type_count_cust, custodian_name]
								batch.insert(insert_data)

								system_search_term = "item-date:#{dateitemsearch} and kind:system and custodian:'#{custodian_name}'"
								system_type_count_cust = nuix_case.count(system_search_term)
								if system_type_count_cust > 0
									system_total_size += nuix_case.getStatistics.getFileSize(system_search_term)
								end
								system_total_count += system_type_count_cust 
								insert_data = [case_guid, "system", searchdate, system_type_count_cust, custodian_name]
								batch.insert(insert_data)

								nodata_search_term = "item-date:#{dateitemsearch} and kind:no-data and custodian:'#{custodian_name}'"
								nodata_type_count_cust = nuix_case.count(nodata_search_term)
								if nodata_type_count_cust > 0
									nodata_total_size += nuix_case.getStatistics.getFileSize(nodata_search_term)
								end
								nodata_total_count += nodata_type_count_cust 
								insert_data = [case_guid, "nodata", searchdate, nodata_type_count_cust, custodian_name]
								batch.insert(insert_data)

								unrecognised_search_term = "item-date:#{dateitemsearch} and kind:unrecognised and custodian:'#{custodian_name}'"
								unrecognised_type_count_cust = nuix_case.count(unrecognised_search_term)
								if unrecognised_type_count_cust > 0
									unrecognised_total_size += nuix_case.getStatistics.getFileSize(unrecognised_search_term)
								end
								unrecognised_total_count = unrecognised_type_count_cust 
								insert_data = [case_guid, "unrecognised", searchdate, unrecognised_type_count_cust, custodian_name]
								batch.insert(insert_data)

								log_search_term = "item-date:#{dateitemsearch} and kind:log and custodian:'#{custodian_name}'"
								log_type_count_cust = nuix_case.count(log_search_term)
								if log_type_count_cust > 0
									log_total_size += nuix_case.getStatistics.getFileSize(log_search_term)
								end
								log_total_count = log_type_count_cust 
								insert_data = [case_guid, "log", searchdate, log_type_count_cust, custodian_name]
								batch.insert(insert_data)
							else
								insert_data = [case_guid, "email",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
								insert_data = [case_guid, "calendar",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
								insert_data = [case_guid, "contact",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
								insert_data = [case_guid, "Documents",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
								insert_data = [case_guid, "Spreadsheets",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
								insert_data = [case_guid, "Presentations",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
								insert_data = [case_guid, "Image",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
								insert_data = [case_guid, "Drawings",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
								insert_data = [case_guid, "Other Documents",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
								insert_data = [case_guid, "Multimedia",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
								insert_data = [case_guid, "Database",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
								insert_data = [case_guid, "Container",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
								insert_data = [case_guid, "System",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
								insert_data = [case_guid, "No Data",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
								insert_data = [case_guid, "Unrecognised",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
								insert_data = [case_guid, "Logs",searchdate, 0,"#{custodian_name}"]
								batch.insert(insert_data)
							end
						end
						searchend = DateTime.now
						searchendtime = searchend.strftime("%d/%m/%Y %H:%M")
						elapsed_seconds = ((searchend - searchstart) * 24 * 60 * 60).to_i
						puts "Search duration was - #{elapsed_seconds}"
					else
						counter += 1
						puts "#{counter} - There is NO data for this Date - #{searchdate}"
						insert_data = [case_guid, "email",searchdate, 0,"NA"]
						batch.insert(insert_data)
						insert_data = [case_guid, "calendar",searchdate, 0,"NA"]
						batch.insert(insert_data)
						insert_data = [case_guid, "contact",searchdate, 0,"NA"]
						batch.insert(insert_data)
						insert_data = [case_guid, "Documents",searchdate, 0,"NA"]
						batch.insert(insert_data)
						insert_data = [case_guid, "Spreadsheets",searchdate, 0,"NA"]
						batch.insert(insert_data)
						insert_data = [case_guid, "Presentations",searchdate, 0,"NA"]
						batch.insert(insert_data)
						insert_data = [case_guid, "Image",searchdate, 0,"NA"]
						batch.insert(insert_data)
						insert_data = [case_guid, "Drawings",searchdate, 0,"NA"]
						batch.insert(insert_data)
						insert_data = [case_guid, "Other Documents",searchdate, 0,"NA"]
						batch.insert(insert_data)
						insert_data = [case_guid, "Multimedia",searchdate, 0,"NA"]
						batch.insert(insert_data)
						insert_data = [case_guid, "Database",searchdate, 0,"NA"]
						batch.insert(insert_data)
						insert_data = [case_guid, "Container",searchdate, 0,"NA"]
						batch.insert(insert_data)
						insert_data = [case_guid, "System",searchdate, 0,"NA"]
						batch.insert(insert_data)
						insert_data = [case_guid, "No Data",searchdate, 0,"NA"]
						batch.insert(insert_data)
						insert_data = [case_guid, "Unrecognised",searchdate, 0,"NA"]
						batch.insert(insert_data)
						insert_data = [case_guid, "Logs",searchdate, 0,"NA"]
						batch.insert(insert_data)
					end
				end
			end
			kinds_count = kinds_count + 'Email::' + "#{email_total_count}".to_s + "::#{email_total_size}" + ' ;'
			kinds_count = kinds_count + 'Calendar::' + "#{calendar_total_count}".to_s + "::#{calendar_total_size}" + ';'
			kinds_count = kinds_count + 'Contact::' + "#{contact_total_count}".to_s + "::#{contact_total_size}" + ';'
			kinds_count = kinds_count + 'Documents::' + "#{documents_total_count}".to_s+ "::#{documents_total_size}" + ';'
			kinds_count = kinds_count + 'Spreadsheets::' + "#{spreadsheets_total_count}".to_s  + "::#{spreadsheets_total_size}" + ';'
			kinds_count = kinds_count + 'Presentations::' + "#{presentations_total_count}".to_s + "::#{presentations_total_size}" + ';'
			kinds_count = kinds_count + 'Image::' + "#{image_total_count}".to_s + "::#{image_total_size}" + ';'
			kinds_count = kinds_count + 'Drawings::' + "#{drawings_total_count}".to_s + "::#{drawings_total_size}" + ';'
			kinds_count = kinds_count + 'Other Documents::' + "#{otherdocuments_total_count}".to_s + "::#{otherdocuments_total_size}" + ';'
			kinds_count = kinds_count + 'Multimedia::' + "#{multimedia_total_count}".to_s + "::#{multimedia_total_size}" + ';'
			kinds_count = kinds_count + 'Database::' + "#{database_total_count}".to_s + "::#{database_total_size}" + ';'
			kinds_count = kinds_count + 'Container::' + "#{container_total_count}".to_s + "::#{container_total_size}" + ';'
			kinds_count = kinds_count + 'System::' + "#{system_total_count}".to_s + "::#{system_total_size}" + ';'
			kinds_count = kinds_count + 'No Data::' + "#{nodata_total_count}".to_s + "::#{nodata_total_size}" + ';'
			kinds_count = kinds_count + 'Unrecognised::' + "#{unrecognised_total_count}".to_s + "::#{unrecognised_total_size}" + ';'
			kinds_count = kinds_count + 'Logs::' + "#{log_total_count}".to_s + "::#{log_total_size}" + ';'
			# insert all the data collected into the database for the case up to the point 
			case_kinds_data = [kinds_count, 70]
			db.update("UPDATE NuixReportingInfo SET ItemTypes = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", case_kinds_data)
		rescue StandardError => msg
			puts "Error in getDateRangeInfo + #{msg}"
		ensure

		end
	end

	# get all the irregular items from the case
	def self.getIrregularItems(nuix_case, case_guid, showsizein, decimal_places, db)
		begin
			corrupted_container_count = nuix_case.count('properties:FailureDetail AND NOT flag:encrypted AND has-text:0 AND ( has-embedded-data:1 OR kind:container OR kind:database )')
			if corrupted_container_count == 0 
				corrupted_container_size = 0
			else
				corrupted_container_size = nuix_case.getStatistics.getFileSize('properties:FailureDetail AND NOT flag:encrypted AND has-text:0 AND ( has-embedded-data:1 OR kind:container OR kind:database )')
			end
			unsupported_container_count = nuix_case.count('kind:( container OR database ) AND NOT flag:encrypted AND has-embedded-data:0 AND NOT flag:partially_processed AND NOT flag:not_processed AND NOT properties:FailureDetail')
			if unsupported_container_count  == 0 
				unsupported_container_size = 0
			else
				unsupported_container_size = nuix_case.getStatistics.getFileSize('kind:( container OR database ) AND NOT flag:encrypted AND has-embedded-data:0 AND NOT flag:partially_processed AND NOT flag:not_processed AND NOT properties:FailureDetail')
			end
			nonsearchable_pdfs_count = nuix_case.count('mime-type:application/pdf AND NOT content:*')
			if nonsearchable_pdfs_count == 0 
				nonsearchable_pdfs_size = 0
			else
				nonsearchable_pdfs_size = nuix_case.getStatistics.getFileSize('mime-type:application/pdf AND NOT content:*')
			end
			text_updated_count = nuix_case.count('modifications:text_updated')
			if text_updated_count == 0 
				text_updated_size = 0
			else
				text_updated_size = nuix_case.getStatistics.getFileSize('modifications:text_updated')
			end
			bad_extension_count = nuix_case.count('flag:irregular_file_extension')
			if bad_extension_count == 0 
				bad_extension_size = 0
			else
				bad_extension_size = nuix_case.getStatistics.getFileSize('flag:irregular_file_extension')
			end
			unrecognized_count = nuix_case.count('kind:unrecognised')
			if unrecognized_count == 0 
				unrecognized_size = 0
			else
				unrecognized_size = nuix_case.getStatistics.getFileSize('kind:unrecognised')
			end
			unsupported_count = nuix_case.count('NOT flag:encrypted AND has-embedded-data:0 AND ( ( has-text:0 AND has-image:0 AND NOT flag:not_processed AND NOT kind:multimedia AND NOT mime-type:application/vnd.ms-shortcut AND NOT mime-type:application/x-contact AND NOT kind:system AND NOT mime-type:( application/vnd.apache-error-log-entry OR application/vnd.linux-syslog-entry OR application/vnd.logstash-log-entry OR application/vnd.ms-iis-log-entry OR application/vnd.ms-windows-event-log-record OR application/vnd.ms-windows-event-logx-record OR application/vnd.ms-windows-setup-api-win7-win8-log-boot-entry OR application/vnd.ms-windows-setup-api-win7-win8-log-section-entry OR application/vnd.ms-windows-setup-api-xp-log-entry OR application/vnd.squid-access-log-entry OR application/vnd.tcpdump.record OR application/vnd.tcpdump.tcp.stream OR application/vnd.tcpdump.udp.stream OR application/x-pcapng-entry OR filesystem/x-linux-login-logfile-record OR filesystem/x-ntfs-logfile-record OR server/dropbox-log-event OR text/x-common-log-entry OR text/x-log-entry ) AND NOT kind:log AND NOT mime-type:application/vnd.ms-exchange-stm ) OR mime-type:application/vnd.lotus-notes )')
			if unsupported_count == 0 
				unsupported_size = 0
			else
				unsupported_size = nuix_case.getStatistics.getFileSize('NOT flag:encrypted AND has-embedded-data:0 AND ( ( has-text:0 AND has-image:0 AND NOT flag:not_processed AND NOT kind:multimedia AND NOT mime-type:application/vnd.ms-shortcut AND NOT mime-type:application/x-contact AND NOT kind:system AND NOT mime-type:( application/vnd.apache-error-log-entry OR application/vnd.linux-syslog-entry OR application/vnd.logstash-log-entry OR application/vnd.ms-iis-log-entry OR application/vnd.ms-windows-event-log-record OR application/vnd.ms-windows-event-logx-record OR application/vnd.ms-windows-setup-api-win7-win8-log-boot-entry OR application/vnd.ms-windows-setup-api-win7-win8-log-section-entry OR application/vnd.ms-windows-setup-api-xp-log-entry OR application/vnd.squid-access-log-entry OR application/vnd.tcpdump.record OR application/vnd.tcpdump.tcp.stream OR application/vnd.tcpdump.udp.stream OR application/x-pcapng-entry OR filesystem/x-linux-login-logfile-record OR filesystem/x-ntfs-logfile-record OR server/dropbox-log-event OR text/x-common-log-entry OR text/x-log-entry ) AND NOT kind:log AND NOT mime-type:application/vnd.ms-exchange-stm ) OR mime-type:application/vnd.lotus-notes )')
			end
			empty_count = nuix_case.count('mime-type:application/x-empty')
			if empty_count == 0 
				empty_size = 0
			else
				empty_size = nuix_case.getStatistics.getFileSize('mime-type:application/x-empty')
			end
			encrypted_count = nuix_case.count('flag:encrypted')
			if encrypted_count == 0 
				encrypted_size = 0
			else
				encrypted_size = nuix_case.getStatistics.getFileSize('mime-type:application/x-empty')
			end
			decrypted_count = nuix_case.count('flag:decrypted')
			if decrypted_count == 0 
				decrypted_size = 0
			else
				decrypted_size = nuix_case.getStatistics.getFileSize('flag:decrypted')
			end
			deleted_count = nuix_case.count('flag:deleted')
			if deleted_count == 0 
				deleted_size = 0
			else
				deleted_size = nuix_case.getStatistics.getFileSize('flag:deleted')
			end
			corrupted_count = nuix_case.count('properties:FailureDetail AND NOT flag:encrypted')
			if corrupted_count == 0 
				corrupted_size = 0
			else
				corrupted_size = nuix_case.getStatistics.getFileSize('properties:FailureDetail AND NOT flag:encrypted')
			end
				digest_mismatch_count = nuix_case.count('flag:digest_mismatch')
			if digest_mismatch_count  == 0 
				digest_mismatch_size = 0
			else
				digest_mismatch_size = nuix_case.getStatistics.getFileSize('flag:digest_mismatch')
			end
			text_stripped_count = nuix_case.count('flag:text_stripped')
			if text_stripped_count == 0 
				text_stripped_size = 0
			else
				text_stripped_size = nuix_case.getStatistics.getFileSize('flag:text_stripped')
			end
			text_not_indexed_count = nuix_case.count('flag:text_not_indexed')
			if text_not_indexed_count == 0 
				text_not_indexed_size = 0
			else
				text_not_indexed_size = nuix_case.getStatistics.getFileSize('flag:text_not_indexed')
			end
			license_restricted_count = nuix_case.count('flag:licence_restricted')
			if license_restricted_count == 0 
				license_restricted_size = 0
			else
				license_restricted_size = nuix_case.getStatistics.getFileSize('flag:licence_restricted')
			end
			not_processed_count = nuix_case.count('flag:not_processed')
			if not_processed_count == 0 
				not_processed_size = 0
			else
				not_processed_size = nuix_case.getStatistics.getFileSize('flag:not_processed')
			end
			partially_processed_count = nuix_case.count('flag:partially_processed')
			if partially_processed_count == 0 
				partially_processed_size = 0
			else
				partially_processed_size = nuix_case.getStatistics.getFileSize('flag:partially_processed')
			end
			text_not_processed_count = nuix_case.count('flag:text_not_processed')
			if text_not_processed_count == 0 
				text_not_processed_size = 0
			else
				text_not_processed_size = nuix_case.getStatistics.getFileSize('flag:text_not_processed')
			end
			images_not_processed_count = nuix_case.count('flag:images_not_processed')
			if images_not_processed_count == 0 
				images_not_processed_size = 0
			else
				images_not_processed_size = nuix_case.getStatistics.getFileSize('flag:images_not_processed')
			end
			reloaded_count = nuix_case.count('flag:reloaded')
			if reloaded_count == 0 
				reloaded_size = 0
			else
				reloaded_size = nuix_case.getStatistics.getFileSize('flag:reloaded')
			end
			poisoned_count = nuix_case.count('flag:poison')
			if poisoned_count == 0 
				poisoned_size = 0
			else
				poisoned_size = nuix_case.getStatistics.getFileSize('flag:poison')
			end
			slack_space_count = nuix_case.count('flag:slack_space')
			if slack_space_count == 0 
				slack_space_size = 0
			else
				slack_space_size = nuix_case.getStatistics.getFileSize('flag:slack_space')
			end
			unallocated_space_count = nuix_case.count('flag:unallocated_space')
			if unallocated_space_count == 0 
				unallocated_space_size = 0
			else
				unallocated_space_size = nuix_case.getStatistics.getFileSize('flag:unallocated_space')
			end
			manually_added_count = nuix_case.count('flag:manually_added')
			if manually_added_count == 0 
				manually_added_size = 0
			else
				manually_added_size = nuix_case.getStatistics.getFileSize('flag:manually_added')
			end
			carved_count = nuix_case.count('flag:carved')
			if carved_count == 0 
				carved_size = 0
			else
				carved_size = nuix_case.getStatistics.getFileSize('flag:carved')
			end
			fully_recovered_count = nuix_case.count('flag:fully_recovered')
			if fully_recovered_count == 0 
				fully_recovered_size = 0
			else
				fully_recovered_size = nuix_case.getStatistics.getFileSize('flag:fully_recovered')
			end
			partially_recovered_count = nuix_case.count('flag:partially_recovered')
			if partially_recovered_count == 0 
				partially_recovered_size = 0
			else
				partially_recovered_size = nuix_case.getStatistics.getFileSize('flag:partially_recovered')
			end
			metadata_recovered_count = nuix_case.count('flag:metadata_recovered')
			if metadata_recovered_count == 0 
				metadata_recovered_size = 0
			else
				metadata_recovered_size = nuix_case.getStatistics.getFileSize('flag:metadata_recovered')
			end
			hidden_stream_count = nuix_case.count('flag:hidden_stream')
			if hidden_stream_count == 0 
				hidden_stream_size = 0
			else
				hidden_stream_size = nuix_case.getStatistics.getFileSize('flag:hidden_stream')
			end
			case showsizein
				when "byte"
					corrupted_container_size = corrupted_container_size  
					unsupported_container_size = unsupported_container_size 
					nonsearchable_pdfs_size = nonsearchable_pdfs_size  
					text_updated_size = text_updated_size 
					bad_extension_size = bad_extension_size 
					unsupported_size = unsupported_size 
					unrecognized_size = unrecognized_size 
					empty_size = empty_size 
					encrypted_size = encrypted_size 
					decrypted_size = decrypted_size 
					deleted_size = deleted_size 
					corrupted_size = corrupted_size 
					digest_mismatch_size = digest_mismatch_size 
					text_stripped_size = text_stripped_size 
					text_not_indexed_size = text_not_indexed_size 
					license_restricted_size = license_restricted_size 
					not_processed_size = not_processed_size 
					partially_processed_size = partially_processed_size 
					text_not_processed_size = text_not_processed_size 
					images_not_processed_size = images_not_processed_size 
					reloaded_size = reloaded_size 
					# poisoned_size = poisoned_size 
					slack_space_size = slack_space_size 
					unallocated_space_size = unallocated_space_size 
					manually_added_size = manually_added_size 
					carved_size = carved_size 
					fully_recovered_size = fully_recovered_size 
					partially_recovered_size = partially_recovered_size 
					metadata_recovered_size = metadata_recovered_size 
					hidden_stream_size = hidden_stream_size 
				when "kb"
					corrupted_container_size = corrupted_container_size.to_kb(decimal_places)
					unsupported_container_size = unsupported_container_size.to_kb(decimal_places)
					nonsearchable_pdfs_size = nonsearchable_pdfs_size.to_kb(decimal_places)
					text_updated_size = text_updated_size.to_kb(decimal_places)
					bad_extension_size = bad_extension_size.to_kb(decimal_places)
					unsupported_size = unsupported_size.to_kb(decimal_places)
					unrecognized_size = unrecognized_size.to_kb(decimal_places)
					empty_size = empty_size.to_kb(decimal_places)
					encrypted_size = encrypted_size.to_kb(decimal_places)
					decrypted_size = decrypted_size.to_kb(decimal_places)
					deleted_size = deleted_size.to_kb(decimal_places)
					corrupted_size = corrupted_size.to_kb(decimal_places)
					digest_mismatch_size = digest_mismatch_size.to_kb(decimal_places)
					text_stripped_size = text_stripped_size.to_kb(decimal_places)
					text_not_indexed_size = text_not_indexed_size.to_kb(decimal_places)
					license_restricted_size = license_restricted_size.to_kb(decimal_places)
					not_processed_size = not_processed_size.to_kb(decimal_places)
					partially_processed_size = partially_processed_size.to_kb(decimal_places)
					text_not_processed_size = text_not_processed_size.to_kb(decimal_places)
					images_not_processed_size = images_not_processed_size.to_kb(decimal_places)
					reloaded_size = reloaded_size.to_kb(decimal_places)
					poisoned_size = poisoned_size.to_kb(decimal_places)
					slack_space_size = slack_space_size.to_kb(decimal_places)
					unallocated_space_size = unallocated_space_size.to_kb(decimal_places)
					manually_added_size = manually_added_size.to_kb(decimal_places)
					carved_size = carved_size.to_kb(decimal_places)
					fully_recovered_size = fully_recovered_size.to_kb(decimal_places)
					partially_recovered_size = partially_recovered_size.to_kb(decimal_places)
					metadata_recovered_size = metadata_recovered_size.to_kb(decimal_places)
					hidden_stream_size = hidden_stream_size.to_kb(decimal_places)
				when "mb"
					corrupted_container_size = corrupted_container_size.to_mb(decimal_places)
					unsupported_container_size = unsupported_container_size.to_mb(decimal_places)
					nonsearchable_pdfs_size = nonsearchable_pdfs_size.to_mb(decimal_places)
					text_updated_size = text_updated_size.to_mb(decimal_places)
					bad_extension_size = bad_extension_size.to_mb(decimal_places)
					unsupported_size = unsupported_size.to_mb(decimal_places)
					unrecognized_size = unrecognized_size.to_mb(decimal_places)
					empty_size = empty_size.to_mb(decimal_places)
					encrypted_size = encrypted_size.to_mb(decimal_places)
					decrypted_size = decrypted_size.to_mb(decimal_places)
					deleted_size = deleted_size.to_mb(decimal_places)
					corrupted_size = corrupted_size.to_mb(decimal_places)
					digest_mismatch_size = digest_mismatch_size.to_mb(decimal_places)
					text_stripped_size = text_stripped_size.to_mb(decimal_places)
					text_not_indexed_size = text_not_indexed_size.to_mb(decimal_places)
					license_restricted_size = license_restricted_size.to_mb(decimal_places)
					not_processed_size = not_processed_size.to_mb(decimal_places)
					partially_processed_size = partially_processed_size.to_mb(decimal_places)
					text_not_processed_size = text_not_processed_size.to_mb(decimal_places)
					images_not_processed_size = images_not_processed_size.to_mb(decimal_places)
					reloaded_size = reloaded_size.to_mb(decimal_places)
					poisoned_size = poisoned_size.to_mb(decimal_places)
					slack_space_size = slack_space_size.to_mb(decimal_places)
					unallocated_space_size = unallocated_space_size.to_mb(decimal_places)
					manually_added_size = manually_added_size.to_mb(decimal_places)
					carved_size = carved_size.to_mb(decimal_places)
					fully_recovered_size = fully_recovered_size.to_mb(decimal_places)
					partially_recovered_size = partially_recovered_size.to_mb(decimal_places)
					metadata_recovered_size = metadata_recovered_size.to_mb(decimal_places)
					hidden_stream_size = hidden_stream_size.to_mb(decimal_places)
				when "gb"
					corrupted_container_size = corrupted_container_size.to_gb(decimal_places)
					unsupported_container_size = unsupported_container_size.to_gb(decimal_places)
					nonsearchable_pdfs_size = nonsearchable_pdfs_size.to_gb(decimal_places)
					text_updated_size = text_updated_size.to_gb(decimal_places)
					bad_extension_size = bad_extension_size.to_gb(decimal_places)
					unsupported_size = unsupported_size.to_gb(decimal_places)
					unrecognized_size = unrecognized_size.to_gb(decimal_places)
					empty_size = empty_size.to_gb(decimal_places)
					encrypted_size = encrypted_size.to_gb(decimal_places)
					decrypted_size = decrypted_size.to_gb(decimal_places)
					deleted_size = deleted_size.to_gb(decimal_places)
					corrupted_size = corrupted_size.to_gb(decimal_places)
					digest_mismatch_size = digest_mismatch_size.to_gb(decimal_places)
					text_stripped_size = text_stripped_size.to_gb(decimal_places)
					text_not_indexed_size = text_not_indexed_size.to_gb(decimal_places)
					license_restricted_size = license_restricted_size.to_gb(decimal_places)
					not_processed_size = not_processed_size.to_gb(decimal_places)
					partially_processed_size = partially_processed_size.to_gb(decimal_places)
					text_not_processed_size = text_not_processed_size.to_gb(decimal_places)
					images_not_processed_size = images_not_processed_size.to_gb(decimal_places)
					reloaded_size = reloaded_size.to_gb(decimal_places)
					poisoned_size = poisoned_size.to_gb(decimal_places)
					slack_space_size = slack_space_size.to_gb(decimal_places)
					unallocated_space_size = unallocated_space_size.to_gb(decimal_places)
					manually_added_size = manually_added_size.to_gb(decimal_places)
					carved_size = carved_size.to_gb(decimal_places)
					fully_recovered_size = fully_recovered_size.to_gb(decimal_places)
					partially_recovered_size = partially_recovered_size.to_gb(decimal_places)
					metadata_recovered_size = metadata_recovered_size.to_gb(decimal_places)
					hidden_stream_size = hidden_stream_size.to_gb(decimal_places)
				else
					corrupted_container_size = corrupted_container_size 
					unsupported_container_size = unsupported_container_size 
					nonsearchable_pdfs_size = nonsearchable_pdfs_size  
					text_updated_size = text_updated_size 
					bad_extension_size = bad_extension_size 
					unsupported_size = unsupported_size 
					unrecognized_size = unrecognized_size 
					empty_size = empty_size 
					encrypted_size = encrypted_size 
					decrypted_size = decrypted_size 
					deleted_size = deleted_size 
					corrupted_size = corrupted_size 
					digest_mismatch_size = digest_mismatch_size 
					text_stripped_size = text_stripped_size 
					text_not_indexed_size = text_not_indexed_size 
					license_restricted_size = license_restricted_size 
					not_processed_size = not_processed_size 
					partially_processed_size = partially_processed_size 
					text_not_processed_size = text_not_processed_size 
					images_not_processed_size = images_not_processed_size 
					reloaded_size = reloaded_size 
					poisoned_size = poisoned_size 
					slack_space_size = slack_space_size 
					unallocated_space_size = unallocated_space_size 
					manually_added_size = manually_added_size 
					carved_size = carved_size 
					fully_recovered_size = fully_recovered_size 
					partially_recovered_size = partially_recovered_size 
					metadata_recovered_size = metadata_recovered_size 
					hidden_stream_size = hidden_stream_size 
			end
			
			irregular_items_count = 'Corrupted_container::' + "#{corrupted_container_count}".to_s + "::#{corrupted_container_size};"
			irregular_items_count = irregular_items_count + 'Unsupported_Container::' + "#{unsupported_container_count}".to_s + "::#{unsupported_container_size};"
			irregular_items_count = irregular_items_count + 'Nonsearchable_PDFs::' + "#{nonsearchable_pdfs_count}".to_s + "::#{nonsearchable_pdfs_size};"
			irregular_items_count = irregular_items_count + 'Text_updated::' + "#{text_updated_count}".to_s + "::#{text_updated_size};"
			irregular_items_count = irregular_items_count + 'Bad_extension::' + "#{bad_extension_count}".to_s + "::#{bad_extension_size};"
			irregular_items_count = irregular_items_count + 'Unrecognized::' + "#{unrecognized_count}".to_s + "::#{unrecognized_size};"
			irregular_items_count = irregular_items_count + 'Unsupported::' + "#{unsupported_count}".to_s + "::#{unsupported_size};"
			irregular_items_count = irregular_items_count + 'Empty::' + "#{empty_count}".to_s + "::#{empty_size};"
			irregular_items_count = irregular_items_count + 'Encrypted::' + "#{encrypted_count}".to_s + "::#{encrypted_size};"
			irregular_items_count = irregular_items_count + 'Decrypted::' + "#{decrypted_count}".to_s + "::#{decrypted_size};"
			irregular_items_count = irregular_items_count + 'Deleted::' + "#{deleted_count}".to_s + "::#{deleted_size};"
			irregular_items_count = irregular_items_count + 'Corrupted::' + "#{corrupted_count}".to_s + "::#{corrupted_size};"
			irregular_items_count = irregular_items_count + 'Digest_mismatch::' + "#{digest_mismatch_count}".to_s + "::#{digest_mismatch_size};"
			irregular_items_count = irregular_items_count + 'Text_stripped::' + "#{text_stripped_count}".to_s + "::#{text_stripped_size};"
			irregular_items_count = irregular_items_count + 'Text_Not_Indexed::' + "#{text_not_indexed_count}".to_s + "::#{text_not_indexed_size};"
			irregular_items_count = irregular_items_count + 'License_restricted::' + "#{license_restricted_count}".to_s + "::#{license_restricted_size};"
			irregular_items_count = irregular_items_count + 'Not_Processed::' + "#{not_processed_count}".to_s + "::#{not_processed_size};"
			irregular_items_count = irregular_items_count + 'Partially_processed::' + "#{partially_processed_count}".to_s + "::#{partially_processed_size};"
			irregular_items_count = irregular_items_count + 'Text_Not_Processed::' + "#{text_not_processed_count}".to_s + "::#{text_not_processed_size};"
			irregular_items_count = irregular_items_count + 'Images_Not_Processed::' + "#{images_not_processed_count}".to_s + "::#{images_not_processed_size};"
			irregular_items_count = irregular_items_count + 'Reload::' + "#{reloaded_count}".to_s + "::#{reloaded_size};"
			irregular_items_count = irregular_items_count + 'Poisoned::' + "#{poisoned_count}".to_s + "::#{poisoned_size};"
			irregular_items_count = irregular_items_count + 'Slack_space::' + "#{slack_space_count}".to_s + "::#{slack_space_size};"
			irregular_items_count = irregular_items_count + 'Unallocated_Space::' + "#{unallocated_space_count}".to_s + "::#{unallocated_space_size};"
			irregular_items_count = irregular_items_count + 'Manually_added::' + "#{manually_added_count}".to_s + "::#{manually_added_size};"
			irregular_items_count = irregular_items_count + 'Carved::' + "#{carved_count}".to_s + "::#{carved_size};";
			irregular_items_count = irregular_items_count + 'Fully_Recovered::' + "#{fully_recovered_count}".to_s + "::#{fully_recovered_size};"
			irregular_items_count = irregular_items_count + 'Partially_Recovered::' + "#{partially_recovered_count}".to_s + "::#{partially_recovered_size};"
			irregular_items_count = irregular_items_count + 'Metadata_Recovered::' + "#{metadata_recovered_count}".to_s + "::#{metadata_recovered_size};"
			irregular_items_count = irregular_items_count + 'Hidden_Stream::' + "#{hidden_stream_count}".to_s + "::#{hidden_stream_size};"
			case_irregularitems_data = [irregular_items_count, 90]
			db.update("UPDATE NuixReportingInfo SET IrregularItems = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", case_irregularitems_data)
		rescue StandardError => msg
			puts "Error in getIrregularItems + #{msg}"
		ensure

		end
	end

	# get the item set info from the case
	def self.getItemSetInfo(nuix_case, case_guid, showsizein, decimal_places, begindate, enddate, db)	
		begin
			puts "Getting Item set info for #{begindate} to #{enddate}"
			all_itemset_info=''
			itemsets = nuix_case.getAllItemSets()
			itemsets.each do |itemset|
				itemsetname = itemset.getName
				itemsetguid = itemset.getGuid
				item_set_batches = itemset.getBatches
				sorted_batches = item_set_batches.sort_by{|batch| batch.getCreated}
				sorted_batches.each do |batch|
					batchdate = batch.getCreated
					batchdate = Date.parse(batchdate.to_s.split("T").first)
					if batchdate.between?(begindate,enddate)
						batchname = batch.getName
						item_set_query="item-set-batch:#{itemsetguid}"
						batch_originals_query = "item-set-batch:#{itemsetguid};originals;#{batchname}"
						batch_duplicates_query = "item-set-batch:#{itemsetguid};duplicates;#{batchname}"
						batchoriginalsfilesize = nuix_case.getStatistics.getFileSize(batch_originals_query)
						batchoriginalsauditsize = nuix_case.getStatistics.getAuditSize(batch_originals_query)
						batchduplicatesfilesize = nuix_case.getStatistics.getFileSize(batch_duplicates_query)
						batchduplicatesauditsize = nuix_case.getStatistics.getAuditSize(batch_duplicates_query)
						batchoriginalscount = nuix_case.count(batch_originals_query)
						batchduplicatecount = nuix_case.count(batch_duplicates_query)
						case showsizein
							when "byte"
								batchoriginalsfilesize = batchoriginalsfilesize
								batchoriginalsauditsize = batchoriginalsauditsize
								batchduplicatesfilesize = batchduplicatesfilesize
								batchduplicatesauditsize = batchduplicatesauditsize
							when "kb"
								batchoriginalsfilesize = batchoriginalsfilesize.to_kb(decimal_places)
								batchoriginalsauditsize = batchoriginalsauditsize.to_kb(decimal_places)
								batchduplicatesfilesize = batchduplicatesfilesize.to_kb(decimal_places)
								batchduplicatesauditsize = batchduplicatesauditsize.to_kb(decimal_places)
							when "mb"
								batchoriginalsfilesize = batchoriginalsfilesize.to_mb(decimal_places)
								batchoriginalsauditsize = batchoriginalsauditsize.to_mb(decimal_places)
								batchduplicatesfilesize = batchduplicatesfilesize.to_mb(decimal_places)
								batchduplicatesauditsize = batchduplicatesauditsize.to_mb(decimal_places)
							when "gb"
								batchoriginalsfilesize = batchoriginalsfilesize.to_gb(decimal_places)
								batchoriginalsauditsize = batchoriginalsauditsize.to_gb(decimal_places)
								batchduplicatesfilesize = batchduplicatesfilesize.to_gb(decimal_places)
								batchduplicatesauditsize = batchduplicatesauditsize.to_gb(decimal_places)
							else
								batchoriginalsfilesize = batchoriginalsfilesize
								batchoriginalsauditsize = batchoriginalsauditsize
								batchduplicatesfilesize = batchduplicatesfilesize
								batchduplicatesauditsize = batchduplicatesauditsize
						end

						all_itemset_info = all_itemset_info + "#{itemsetname}-#{batchname}" + '::' + "#{batchdate}" + '::' + "#{batchoriginalscount}".to_s + '::' + "#{batchoriginalsfilesize}".to_s + '::' + "#{batchoriginalsauditsize}".to_s + '::' + "#{batchduplicatecount}".to_s + '::' + "#{batchduplicatesfilesize}".to_s + '::' + "#{batchduplicatesauditsize}".to_s + ";"
					end
				end
			end
				case_itemset_info = [all_itemset_info, 20]
				puts "Updating Percent Complete to 20 percent for '#{case_guid}' with '#{all_itemset_info}'"
				db.update("UPDATE NuixReportingInfo SET ItemSets = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", case_itemset_info)
		rescue StandardError => msg
			
		   puts "Error in getItemSetInfo + #{msg}"
		ensure
		end
	end

	# if necessary get the search criteria 
	def self.getSearchCriteria(nuix_case, case_guid, searchterm, showsizein, decimal_places, db)
		begin
			userentered_all_custodians_info=''
			userentered_search_hit_count=0
			userentered_search_file_size=0
			userentered_search_criteria = "#{searchterm}"
			userentered_search_hit_count = nuix_case.count(userentered_search_criteria)
			userentered_search_file_size = nuix_case.getStatistics.getFileSize(userentered_search_criteria)
			case showsizein
				when "byte"
					userentered_search_file_size = userentered_search_file_size
				when "kb"
					userentered_search_file_size = userentered_search_file_size.to_kb(decimal_places)
				when "mb"
					userentered_search_file_size = userentered_search_file_size.to_mb(decimal_places)
				when "gb"
					userentered_search_file_size = userentered_search_file_size.to_gb(decimal_places)
				else
					userentered_search_file_size = userentered_search_file_size
			end
			userentered_custodian_name_count=0
			custodian_names = nuix_case.getAllCustodians
			custodian_names.each do |custodian_name|
				search_criteria = "custodian:'#{custodian_name}' and #{searchterm}"
puts "Search Criteria = #{search_criteria}"
				userentered_custodian_name_count = nuix_case.count("#{search_criteria}")
puts "Search Criteria count = #{userentered_custodian_name_count}"
				userentered_all_custodians_info = userentered_all_custodians_info + "#{custodian_name}::#{userentered_custodian_name_count};"
			end
		   case_items_data = [userentered_search_criteria,userentered_all_custodians_info, userentered_search_hit_count, userentered_search_file_size, 85]
		   db.update("UPDATE NuixReportingInfo SET SearchTerm = ?, CustodianSearchHit = ?, HitCount = ?, SearchSize = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", case_items_data)
		rescue StandardError => msg
			puts "Error in getSearchCriteria + #{msg}"
		ensure

		end
	end
	# open the search file and loop through each item in the file if necessary
	def self.getSearchFileCriteria(nuix_case, case_guid, searchtermfile, showsizein, decimal_places, db)
		begin
			userentered_all_custodians_info=''
			userentered_search_hit_count=0
			userentered_search_file_size=0
			db.update("Delete from UCRTSearchTermResults where CaseGUID = '#{case_guid}'")
			sSearchTerm = ''
			sExportFolder = ''
			CSV.foreach("#{searchtermfile}") do |row|
				sSearchTerm = row[0]
				sExportFolder =  row[1]
				userentered_search_hit_count = nuix_case.count(sSearchTerm)
				userentered_search_file_size = nuix_case.getStatistics.getFileSize(sSearchTerm)
				user_search_details = [case_guid, sSearchTerm, userentered_search_hit_count]
				db.update("INSERT INTO UCRTSearchTermResults (CaseGUID, SearchTerm, ItemCount) VALUES ( ?, ?, ?)", user_search_details)
				userentered_custodian_name_count=0
				custodian_names = nuix_case.getAllCustodians
				custodian_names.each do |custodian_name|
					search_criteria = "custodian:\"#{custodian_name}\"" + " and " + "\"#{sSearchTerm}\""
					userentered_custodian_name_count = nuix_case.count("#{search_criteria}")
					userentered_all_custodians_info = userentered_all_custodians_info + "#{custodian_name}" + '::' + "#{userentered_custodian_name_count}".to_s + ";"
					user_search_details = [case_guid, search_criteria, userentered_search_hit_count, custodian_name]
					db.update("INSERT INTO UCRTSearchTermResults (CaseGUID, SearchTerm, Custodian, CustodianSearchTermItemCount) VALUES ( ?, ?, ?, ?)", user_search_details)
				end
			end
			case_items_data = [userentered_all_custodians_info, userentered_search_hit_count, userentered_search_file_size, 85]
			db.update("UPDATE NuixReportingInfo SET CustodianSearchHit = ?, HitCount = ?, SearchSize = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", case_items_data)
		rescue StandardError => msg
			puts "Error in getSearchCriteria + #{msg}"
		ensure

		end
	end

	#export the items in the search if necessary
	def self.nuixExportSearchTermResults(nuix_case, case_guid, searchterm, exportfolder, exportType, showsizein, decimal_places, db)
		items = nuix_case.search(searchterm)
		rightnow = DateTime.now.to_s
		rightnow = rightnow.delete(':')
		batchexportname = "'#{exportfolder}' +  '#{rightnow}' + '\\' + '#{case_name}'"
		exporter = $utilities.createBatchExporter(batchexportname)
		exporter.setParallelProcessingSettings({
			:workerCount => exportWorkers,
			:workerMemory => exportWorkerMemory,
			:workerTemp => "C:\\Temp",
			:embedBroker => true,
			:brokerMemory => 768
		})

		if exportType == "Native - GUID"
			natives_settings = {
				:naming => "guid",
				:path => "NATIVE",
				:mailFormat => "pst",
				:includeAttachments => true,
			}
			exporter.addProduct("native", natives_settings)
			exporter.exportItems(items)
		elsif exportType == "Native - Name"
			natives_settings = {
				:naming => 	"item_name",
				:path => "NATIVE",
				:mailFormat => "pst",
				:includeAttachments => true,
			}
			exporter.addProduct("native", natives_settings)
			exporter.exportItems(items)
		elsif exportType == "PDF"
			pdf_settings = {
				:naming => "guid",
				:path => "PDF",
				:mailFormat => "pst",
				:includeAttachments => true,
			}
			exporter.addProduct("pdf", pdf_settings)
			exporter.exportItems(items)
		elsif exportType == "Mailbox"
			mailboxexporter = utilities.getMailboxExporter()
			natives_settings = {
				:format => "pst",
				:path => nil,
				:failfast => "false"
			}
			directory_name = "#{exportfolder} +  #{rightnow} + \\"
			response = FileUtils.mkdir_p(directory_name)
			exporter.exportItems(items, "#{directory_name} + \\ + case_name + .pst", natives_settings)                            
		elsif exportType == "NLI"
			natives_settings = {
			}
			export_settings = {
			}
			directory_name = "#{exportfolder} +  #{rightnow} + \\"
			response = FileUtils.mkdir_p(directory_name)
			exporter = $utilities.createLogicalImageExporter(directory_name, case_name, export_settings)
			items.each do |item|
				exporter.addItem(item)
			end	
		elsif exportType == "CaseSubset"
			casesubsetexporter = $utilities.getCaseSubsetExporter()
			natives_settings = {
				:naming => "guid",
				:path => "NATIVE",
				:mailFormat => "pst",
				:includeAttachments => true,
			}	
			case_subset_settings = {
				:evidenceStoreCount => 1,
				:includeFamilies => true,
				:copyTags => true,
				:copyComments => true,
				:copyCustodians => true,
				:copyItemSets => true,
				:copyClassifiers => true,
				:copyMarkupSets => true,
				:copyProductionSets => true,
				:copyClusters => true,
				:copyCustomMetadata => true,
				:copyGraphDatabase => false,
				:caseMetadata => nil,
				:processingSettings => nil
			}
			directory_name = '#{exportfolder} +  #{rightnow} + \\'
			response = FileUtils.mkdir_p(directory_name)
			exporter.exportItems(items,'#{directory_name}' +  case_name + '\\' + rightnow  + '\\CaseSubset',case_subset_settings)
		end
	end
	
	# export the items in the search file if necessary
	def self.nuixExportSearchFileResults(nuix_case, case_guid, searchtermfile, exportfolder, exportType, showsizein, decimal_places, db)
		CSV.foreach(searchtermfile) do |row|
			sSearchTerm = row[0]
			sExportFolder = row[1]
			rightnow = DateTime.now.to_s
			rightnow = rightnow.delete(':')
			batchexportname = "'#{exportfolder}' +  '#{rightnow}' + '\\' + '#{case_name}'"
			exporter = $utilities.createBatchExporter(batchexportname)
			exporter.setParallelProcessingSettings({
				:workerCount => exportWorkers,
				:workerMemory => exportWorkerMemory,
				:workerTemp => "C:\\Temp",
				:embedBroker => true,
				:brokerMemory => 768
			})

			if exportType == "Native - GUID"
				natives_settings = {
					:naming => "guid",
					:path => "NATIVE",
					:mailFormat => "pst",
					:includeAttachments => true,
				}
				exporter.addProduct("native", natives_settings)
				exporter.exportItems(items)
			elsif exportType == "Native - Name"
				natives_settings = {
					:naming => 	"item_name",
					:path => "NATIVE",
					:mailFormat => "pst",
					:includeAttachments => true,
				}
				exporter.addProduct("native", natives_settings)
				exporter.exportItems(items)
			elsif exportType == "PDF"
				pdf_settings = {
					:naming => "guid",
					:path => "PDF",
					:mailFormat => "pst",
					:includeAttachments => true,
				}
				exporter.addProduct("pdf", pdf_settings)
				exporter.exportItems(items)
			elsif exportType == "Mailbox"
				mailboxexporter = utilities.getMailboxExporter()
				natives_settings = {
					:format => "pst",
					:path => nil,
					:failfast => "false"
				}
				directory_name = "#{exportfolder} +  #{rightnow} + \\"
				response = FileUtils.mkdir_p(directory_name)
				exporter.exportItems(items, "#{directory_name} + \\ + case_name + .pst", natives_settings)                            
			elsif exportType == "NLI"
				natives_settings = {
				}
				export_settings = {
				}
				directory_name = "#{exportfolder} +  #{rightnow} + \\"
				response = FileUtils.mkdir_p(directory_name)
				exporter = $utilities.createLogicalImageExporter(directory_name, case_name, export_settings)
				items.each do |item|
					exporter.addItem(item)
				end
			elsif exportType == "CaseSubset"
				casesubsetexporter = $utilities.getCaseSubsetExporter()
				natives_settings = {
					:naming => "guid",
					:path => "NATIVE",
					:mailFormat => "pst",
					:includeAttachments => true,
				}
				case_subset_settings = {
					:evidenceStoreCount => 1,
					:includeFamilies => true,
					:copyTags => true,
					:copyComments => true,
					:copyCustodians => true,
					:copyItemSets => true,
					:copyClassifiers => true,
					:copyMarkupSets => true,
					:copyProductionSets => true,
					:copyClusters => true,
					:copyCustomMetadata => true,
					:copyGraphDatabase => false,
					:caseMetadata => nil,
					:processingSettings => nil
				}
				directory_name = '#{exportfolder} +  #{rightnow} + \\'
				response = FileUtils.mkdir_p(directory_name)
				exporter.exportItems(items,'#{directory_name}' +  case_name + '\\' + rightnow  + '\\CaseSubset',case_subset_settings)
			end
		end
	end
end

# function to get the age of a file (how old a given file is
def file_age(name)
	puts "File CTime - " + File.mtime(name).to_s
	(Time.now - File.mtime(name))/(24*3600)
end

# Start the execution of the UCRT 
# before processing no cases have been open so set the case_opened variable to 0
case_opened = 0
# Call the UCRT class getCases method to determine the locations of the cases to search for Nuix cases in 
# The getCases method will return a list of all of the cases in the directories that are specified by the 
# json value "caseslocations" - this value can be a comma separated list of directories
# i.e.   "caseslocations": "C:\\Nuix\\Sample Cases, C:\\Nuix_WORKING\\Cases",
allcases = UCRT.getCases("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getUpgradeValue method to determine if the cases should be upgraded by the script
# The getUpgradeValue will return a true/false value from the json value "upgradecases"
upgradecases = Settings.getUpgradeValue("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getIncludeDateRangeValue method to determine if the script should loop through daily case details
# The getIncludeDateRangeValue will return a true/false value from the json value "includedaterange"
includedaterange = Settings.getIncludeDateRangeValue("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getIgnoreCaseHistoryValue method to determine if the script get the caseHistory
# The getIncludeDateRangeValue will return a true/false value from the json value "ignorecasehistory"
ignorecasehistory = Settings.getIgnoreCaseHistoryValue("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getIncludeAnnotationsValue method to determine if the script get the caseHistory
# The getIncludeDateRangeValue will return a true/false value from the json value "includeannotations"
includeannotations = Settings.getIncludeAnnotationsValue("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getDBLocation method to determine the locations of the sqlite database
dblocation = Settings.getDBlocation("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getExportdirectory method to determine locations of the export directory
exportdirectory = Settings.getExportdirectory("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getExportdirectory method to get the value of the exportfilefile
exportfilename = Settings.getExportFilename("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getExportdirectory method to get the value of the fields to export
# Valid values are "all" or a comma separated list of fields to export from the database when the UCRT has completed
exportcaseinfo = Settings.getExportCaseInfo("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getExportdirectory method to get the value Export Type
# Valid values are "csv", "json", or "xml"
exporttype = Settings.getExporttype("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getUCRTreportinguser method to get the value of ucrtreportinguser - if this value
# is present the UCRT will ignore any history coming from this user as to not give unnecssary statistics for 
# opening and closing the case
ucrtreportinguser = Settings.getUCRTreportinguser("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getReportfrequency method to get the value of getReportFrequency
# Valid values are "daily", "weekly", "monthly", "quarterly", "yearly"
reportfrequency = Settings.getReportfrequency("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getReportfrequency method to get the value of getCSVExportFields
# this is the value of the fields that will be exported in the report
csvexportfields = Settings.getCSVExportFields("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getSearchTermValue method to get the value of "searchterm"
# this is any valid NQL to run against each case 
searchterm = Settings.getSearchTermFileValue("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getSearchTermValue method to get the value of "searchtermfile"
# this is any valid NQL to run against each case - this file can have multiple NQL searches that will run against the cases
searchtermfile = Settings.getSearchTermFileValue("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getCleanupDatabaseValue method to get the value of "cleanupdatabase"
# The getCleanupDatabaseValue will return a true/false value from the json value "cleanupdatabase"
cleanupdatabase = Settings.getCleanupDatabaseValue("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getShowSizeInValue method to get the value of "showsizein"
# value values are "byte", "kb", "mb", "gb"
showsizein = Settings.getShowSizeInValue("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getVersioninfo method to get the value of "nuixversionmapping"
# this will allow the user to map Nuix case versions to the appropriate Nuix console version that can open the cases
versionmapping = Settings.getVersionInfo("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getDecimalPointAccuracy method to get the value of "decimalpointaccuracy"
# this value will allow the user to set how accurate they would like the value after the decimal when reporting
decimal_places = Settings.getDecimalPointAccuracy("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getCleanupFilesValue method to get the value of "cleanupfiles"
# The getCleanupFilesValue will return a true/false value from the json value "cleanupdatabase"
cleanupfiles = Settings.getCleanupFilesValue("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getCleanupFilesRange method to get the value of "cleanupfilesrange"
# The getCleanupFilesRange will return a number for the amount of days in the past to remove files specified in 
# getCleanupFileType in the getCleanupDirectories folders
cleanup_filerange = Settings.getCleanupFileRange("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getCleanupDirectories method to get the value of "cleanupdirectories"
# The getCleanupDirectories will return a comma separated list of directories to clean up files in 
cleanup_directories = Settings.getCleanupDirectories("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getCleanupFileType method to get the value of "cleanupfiletypes"
# The getCleanupFileType will return a comma separated list of types 
cleanup_filetypes = Settings.getCleanupFileType("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getNuixExportSearchResults method to get the value of "nuixexportsearchresults"
# The getNuixExportSearchResults will return a true/false value from the json value "nuixexportsearchresults"
nuixexportsearchresults = Settings.getNuixExportSearchResults("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
# Call the UCRT class getNuixExportType method to get the value of "nuixexporttype"
# The getNuixExportType will return a string - valid values are any valid nuix export type
nuixexporttype = Settings.getNuixExportType("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")

# Get the start time that the UCRT will begin processing
reportloadstart = DateTime.now
# Convert the start time to a the correct date value
reportloadstarttime = reportloadstart.strftime("%m/%d/%Y %H:%M")
# Write the values that were returned from above to the nuix.log log file
puts "Upgrade Cases: #{upgradecases}"
puts "Include Daterange: #{includedaterange}"
puts "Ignore Case History: #{ignorecasehistory}"
puts "Include Annotations: #{includeannotations}"
puts "Database locations: #{dblocation}"
puts "Export Filename: #{exportfilename}"
puts "Database Directory: #{exportdirectory}"
puts "Export Case Info: #{exportcaseinfo}"
puts "Export Type: #{exporttype}"
puts "UCRT Reporting User: #{ucrtreportinguser}"
puts "Report Frequency: #{reportfrequency}"
puts "CSV Export Fields: #{csvexportfields}"
puts "Search Term: #{searchterm}"
puts "Search Term File: #{searchtermfile}"
puts "Cleanup Database: #{cleanupdatabase}"
puts "Show Size In: #{showsizein}"
puts "Version Mapping: #{versionmapping}"
puts "Decimal Places: #{decimal_places}"
puts "Cleanup Files: #{cleanupfiles}"
puts "Cleanup File Range: #{cleanup_filerange}"
puts "Cleanup Directories: #{cleanup_directories}"
puts "Cleanup File Types: #{cleanup_filetypes}"
puts "Nuix Export Search Results: #{nuixexportsearchresults}"
puts "Nuix Export Type: #{nuixexporttype}"

# load the Database.rb_ ruby helper that should be located in the database location
load "#{dblocation}\\Database.rb_"
# load the SQLite.rb_ ruby helper that should be located in the database location
load "#{dblocation}\\SQLite.rb_"
# Create a new instance of the NuixCaseReports.db3 to be used to store the values returned from the methods
db = SQLite.new("#{dblocation}\\NuixCaseReports.db3")

# Get the value for yesterday 
yesterday = Date.today - 1
# Get the current year
currentyear = Date.today.year

begindate = ''
enddate = ''

# If the cleanupdatabase is true then get the GUIDs to clean up from the setting.json file
if cleanupdatabase == "true"
	cleanupcaseguids = UCRT.getCleanupDatabaseGUIDs("C:\\Program Files\\Nuix\\ScriptAutomate\\Settings.json")
	if cleanupcaseguids == "all"
		db.update("Delete from NuixReportingInfo")
		db.update("Delete from UCRTSessionEvents")
		db.update("Delete from UCRTDateRange")
		db.update("Delete from UCRTSearchTermResults")
	else
		db.update("Delete from NuixReportingInfo where CaseGUID in ('#{cleanupcaseguids}')")
		db.update("Delete from UCRTSessionEvents where CaseGUID in ('#{cleanupcaseguids}')")
		db.update("Delete from UCRTDateRange where CaseGUID in ('#{cleanupcaseguids}')")
		db.update("Delete from UCRTSearchTermResults where CaseGUID in ('#{cleanupcaseguids}')")
	end
end

# Get the reportfrequency and set the range appropriately if "daily", "weekly", "monthly", "quarterly", "yearly"
case reportfrequency
	when "daily"
		begindate = Date.today - 1
		enddate = Date.today
	when "Daily"
		begindate = Date.today - 1
		enddate = Date.today
	when "weekly"
		nameofday = Date.today.dayname
		case nameofday
			when "Sunday"
				begindate = Date.today - 7
				enddate = Date.today
			when "Monday"
				begindate = Date.today - 8
				enddate = Date.today - 1
			
			when "Tuesday"
				begindate = Date.today - 9
				enddate = Date.today - 2
			
			when "Wednesday"
				begindate = Date.today - 10
				enddate = Date.today - 3
			
			when "Thursday"
				begindate = Date.today - 11
				enddate = Date.today - 4
			
			when "Friday"
				begindate = Date.today - 12
				enddate = Date.today - 5
			
			when "Saturday"
				begindate = Date.today - 13
				enddate = Date.today - 6
			else
		end				
	when "Weekly"
		nameofday = Date.today.dayname
		case nameofday
			when "Sunday"
				begindate = Date.today - 7
				enddate = Date.today
			when "Monday"
				begindate = Date.today - 8
				enddate = Date.today - 1
			
			when "Tuesday"
				begindate = Date.today - 9
				enddate = Date.today - 2
			
			when "Wednesday"
				begindate = Date.today - 10
				enddate = Date.today - 3
			
			when "Thursday"
				begindate = Date.today - 11
				enddate = Date.today - 4
			
			when "Friday"
				begindate = Date.today - 12
				enddate = Date.today - 5
			
			when "Saturday"
				begindate = Date.today - 13
				enddate = Date.today - 6
			else
		end
	when "monthly"
		monthnumber = Time.now.month

		case monthnumber
			when 1
				begindate = Date.parse('1-1-' + currentyear.to_s)
				enddate = Date.parse('31-1-' + currentyear.to_s)
			when 2
				now = DateTime.now 
				flag = Date.leap?( now.year )
				if flag
					begindate = Date.parse('1-2-' + currentyear.to_s)
					enddate = Date.parse('29-2-' + currentyear.to_s)
				else
					begindate = Date.parse('1-2-' + currentyear.to_s)
					enddate = Date.parse('28-2-' + currentyear.to_s)
				end
			when 3
				begindate = Date.parse('1-3-' + currentyear.to_s)
				enddate = Date.parse('31-3-' + currentyear.to_s)
			
			when 4
				begindate = Date.parse('1-4-' + currentyear.to_s)
				enddate = Date.parse('30-4-' + currentyear.to_s)
			
			when 5
				begindate = Date.parse('1-5-' + currentyear.to_s)
				enddate = Date.parse('31-5-' + currentyear.to_s)
			
			when 6
				begindate = Date.parse('1-6-' + currentyear.to_s)
				enddate = Date.parse('30-6-' + currentyear.to_s)
			
			when 7
				begindate = Date.parse('1-7-' + currentyear.to_s)
				enddate = Date.parse('31-7-' + currentyear.to_s)
			
			when 8
				begindate = Date.parse('1-8-' + currentyear.to_s)
				enddate = Date.parse('31-8-' + currentyear.to_s)
			
			when 9
				begindate = Date.parse('1-9-' + currentyear.to_s)
				enddate = Date.parse('30-9-' + currentyear.to_s)
			
			when 10
				begindate = Date.parse('1-10-' + currentyear.to_s)
				enddate = Date.parse('31-10-' + currentyear.to_s)
			
			when 11
				begindate = Date.parse('1-11-' + currentyear.to_s)
				enddate = Date.parse('30-11-' + currentyear.to_s)
			
			when 12
				begindate = Date.parse('1-12-' + currentyear.to_s)
				enddate = Date.parse('31-12-' + currentyear.to_s)
			else
		end
	when "Monthly"
		monthnumber = Time.now.month

		case monthnumber
			when 1
				begindate = Date.parse('1-1-' + currentyear.to_s)
				enddate = Date.parse('31-1-' + currentyear.to_s)
			when 2
				now = DateTime.now 
				flag = Date.leap?( now.year )
				if flag
					begindate = Date.parse('1-2-' + currentyear.to_s)
					enddate = Date.parse('29-2-' + currentyear.to_s)
				else
					begindate = Date.parse('1-2-' + currentyear.to_s)
					enddate = Date.parse('28-2-' + currentyear.to_s)
				end
			when 3
				begindate = Date.parse('1-3-' + currentyear.to_s)
				enddate = Date.parse('31-3-' + currentyear.to_s)
			
			when 4
				begindate = Date.parse('1-4-' + currentyear.to_s)
				enddate = Date.parse('30-4-' + currentyear.to_s)
			
			when 5
				begindate = Date.parse('1-5-' + currentyear.to_s)
				enddate = Date.parse('31-5-' + currentyear.to_s)
			
			when 6
				begindate = Date.parse('1-6-' + currentyear.to_s)
				enddate = Date.parse('30-6-' + currentyear.to_s)
			
			when 7
				begindate = Date.parse('1-7-' + currentyear.to_s)
				enddate = Date.parse('31-7-' + currentyear.to_s)
			
			when 8
				begindate = Date.parse('1-8-' + currentyear.to_s)
				enddate = Date.parse('31-8-' + currentyear.to_s)
			
			when 9
				begindate = Date.parse('1-9-' + currentyear.to_s)
				enddate = Date.parse('30-9-' + currentyear.to_s)
			
			when 10
				begindate = Date.parse('1-10-' + currentyear.to_s)
				enddate = Date.parse('31-10-' + currentyear.to_s)
			
			when 11
				begindate = Date.parse('1-11-' + currentyear.to_s)
				enddate = Date.parse('30-11-' + currentyear.to_s)
			
			when 12
				begindate = Date.parse('1-12-' + currentyear.to_s)
				enddate = Date.parse('31-12-' + currentyear.to_s)
			else
		end	
	when "quarterly"
		datequarter = yesterday.quarter
		puts datequarter

		case datequarter
			when 1
				begindate = Date.parse('1-1-' + currentyear.to_s)
				enddate = Date.parse('31-3-' + currentyear.to_s)
			when 2
				begindate = Date.parse('1-4-' + currentyear.to_s)
				enddate = Date.parse('30-6-' + currentyear.to_s)
			when 3
				begindate = Date.parse('1-7-' + currentyear.to_s)
				enddate = Date.parse('30-9-' + currentyear.to_s)
			when 4
				begindate = Date.parse('1-10-' + currentyear.to_s)
				enddate = Date.parse('31-12-' + currentyear.to_s)
			else

		end
	when "Quarterly"
		datequarter = yesterday.quarter
		puts datequarter

		case datequarter
			when 1
				begindate = Date.parse('1-1-' + currentyear.to_s)
				enddate = Date.parse('31-3-' + currentyear.to_s)
			when 2
				begindate = Date.parse('1-4-' + currentyear.to_s)
				enddate = Date.parse('30-6-' + currentyear.to_s)
			when 3
				begindate = Date.parse('1-7-' + currentyear.to_s)
				enddate = Date.parse('30-9-' + currentyear.to_s)
			when 4
				begindate = Date.parse('1-10-' + currentyear.to_s)
				enddate = Date.parse('31-12-' + currentyear.to_s)
			else

		end
	when "yearly"
		todaysdate = Date.today
		oneyearago = Date.civil(todaysdate.year-1, todaysdate.month, todaysdate.day)
		begindate = oneyearago
		enddate = todaysdate
	when "Yearly"
		todaysdate = Date.today
		oneyearago = Date.civil(todaysdate.year-1, todaysdate.month, todaysdate.day)
		begindate = oneyearago
		enddate = todaysdate
	when "calendaryear"
		begindate = Date.parse('1-1-' + currentyear.to_s)
		enddate = Date.parse('31-12-' + currentyear.to_s)
	when "Calendaryear"
		begindate = Date.parse('1-1-' + currentyear.to_s)
		enddate = Date.parse('31-12-' + currentyear.to_s)
	when "all"
		begindate = Date.parse('1-1-1990')
		enddate = Date.today
	when "All"
		begindate = Date.parse('1-1-1990')
		enddate = Date.today
	else
end

# Write the begin date and end date to the nuix.log file
puts "Begin date = #{begindate}"
puts "End date = #{enddate}"

# set the value of processed_guids to '' since we have not processed any cases yet
processed_guids = ''

# Loop over every case that is stored in the allcases array
allcases.each do |nuixcase|
	# replace any / with \\ so reporting will be consistent
	nuixcase = nuixcase.gsub("/","\\")
	# set the value for the caseloadstart to be now
	caseloadstart = DateTime.now
	# convert the case load start to the appropriate format
	caseloadstarttime = caseloadstart.strftime("%m/%d/%Y %H:%M")

	# set the counter to 0
	count = 0
	# get the time that the case.fbi2 file was created for reporting purposes
	casecreated = File.ctime(nuixcase)
	# get the time that the case.fbi2 file was modified for reporting purposes
	casemodified = File.mtime(nuixcase)
	# get the folder location of the case.fbi2 file
	casefolder = nuixcase.gsub("\\case.fbi2","")
	# call the UCRT class getCaseDetails method to get the information about the Nuix case that 
	# is stored on disc and not in inside the Nuix case (not necessary to open the case to get this info
	case_details = UCRT.getCaseDetails(casefolder)
	# get the root directory of the case 
	root_dir = java.io.File.new("#{casefolder}")
	# execue the sizeOfDirectory to get the full size of the file that make up the Nuix case for reporting purposes
	disk_size = org.apache.commons.io.FileUtils::sizeOfDirectory(root_dir)
	# depending on the value of show size in - store the values of the disk_size accordingly
	case showsizein
		when "byte"
			disk_size = disk_size
			puts "Disk size in Bytes = #{disk_size}"
		when "kb"
			disk_size = disk_size.to_kb(decimal_places)
			puts "Disk size in KB = #{disk_size}"
		when "mb"
			disk_size = disk_size.to_mb(decimal_places)
			puts "Disk size in MB = #{disk_size}"
		when "gb"
			disk_size = disk_size.to_gb(decimal_places)
			puts "Disk size in GB = #{disk_size}"
		else
			disk_size = disk_size
			puts "Disk size not specified = #{disk_size}"
	end

	# get the case name from the case details 
	case_name = case_details["case_name"]
	# get the case guid from the case details
	case_guid = case_details["case_guid"]
	# get the case type (simple or compound) from the case details
	case_type = case_details["case_type"]
	if case_type == "SIMPLE"
		is_compound = 0
	else
		is_compound = 1
	end
	# get the creation_date from the case details
	creation_date = case_details["creation_date"]
	# get the modified_date from the case details
	modified_date = case_details["modified_date"]
	# get the investigator from the case details
	case_investigator = case_details["investigator"]
	# get the case version from the case details
	case_version = case_details["nuix_version"]
	# get the workerTemp from the case details
	worker_temp = case_details["workerTemp"]
	# get the workerCount from the case details
	worker_count = case_details["workerCount"]
	# get the brokerMemory from the case details
	broker_memory = case_details["brokerMemory"]
	# get the workerMemory from the case details
	worker_memory = case_details["workerMemory"]
	# get the evidence_name from the case details
	evidence_name = case_details["evidence_name"]
	# get the evidence_locations from the case details
	evidence_locations = case_details["evidence_locations"]
	# get the evidence description from the case details
	evidence_description = case_details["evidence_description"]
	# get the is_locked value from the case details
	case_locked = case_details["is_locked"]
	# get the locked_by value from the case details
	case_locked_by = case_details["locked_by"]
	# get the locked_date value from the case details
	case_lock_date = case_details["lock_date"]
	# get the lock_machine value from the case details
	case_lock_machine = case_details["lock_machine"]
	# get the lock_product value from the case details
	case_lock_product = case_details["lock_product"]

	# if the case is locked the update the database and logs accordingly
	if case_locked == "true"
		puts "Case Is Locked - Case #{case_guid} is Locked-#{case_locked_by}-#{case_lock_date}-#{case_lock_machine}"
		count = db.scalar("select count(*) from NuixReportingInfo where CaseGUID = '#{case_guid}'")
		if count == 0
			puts "Inserting into Database with locked case"
			lock_info = "#{case_locked_by}-#{case_lock_machine}-#{case_lock_product}-#{case_lock_date}"
			session_data = [case_guid,'','',case_name,"!!!CASE LOCKED!!!",'0','','','',casefolder,'',case_version,'',disk_size,'0','0',is_compound,'','',case_investigator,'','','',broker_memory,worker_count,worker_memory,evidence_name,evidence_locations,'','','',creation_date,modified_date,'','','','','','','','','','','','','','','','','','','','','','','','','','',lock_info,'False']
			db.update("INSERT INTO NuixReportingInfo(CaseGUID, ReportLoadStart, ReportLoadEnd, CaseName, CollectionStatus, PercentComplete, ReportLoadDuration, BatchLoadInfo, ItemSets, CaseLocation, BackUpLocation, CurrentCaseVersion, UpgradedCaseVersion, CaseSizeOnDisk, CaseFileSize, CaseAuditSize, IsCompound, CasesContained, ContainedInCase, Investigator, InvestigatorSessions, InvalidSessions, InvestigatorTimeSummary, BrokerMemory, WorkerCount, WorkerMemory, EvidenceProcessed, EvidenceLocation, EvidenceCustomMetadata, MimeTypes, ItemTypes, CreationDate, ModifiedDate, LoadDataStart, LoadDataEnd, LoadTime, LoadEvents, TotalLoadTime, ProcessingSpeed, Custodians, CustodianCount, SearchTerm, SearchSize, HitCount, CustodianSearchHit, TotalItemCount, ItemCounts, OriginalItems, DuplicateItems, CaseUsers, ReportLoadTime, NuixLogLocation, OldestItem, NewestItem, Languages, CustomMetadata, CaseDescription, EvidenceDescription, IrregularItems,LockInfo,ReportingInfoCollected) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", session_data)
		else
			puts "Updating Database with locked case"
			collection_status = "!!!CASE LOCKED!!!" 
			lock_info = "#{case_locked_by}-#{case_lock_machine}-#{case_lock_product}-#{case_lock_date}"
			processing_data = [collection_status,lock_info,100]
			db.update("UPDATE NuixReportingInfo SET CollectionStatus = ?, LockInfo = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", processing_data)
		end
		if processed_guids == ''
			processed_guids = processed_guids + "'" + case_guid + "'"
		else
			processed_guids = processed_guids + ",'" + case_guid + "'"
		end
	# if the case is not locked then update the log and database and then open the case to get stats
	else
		puts "Case - #{case_name} - is not locked"
		collection_status = "File System and Case Data collected" 
		processing_data = [collection_status,disk_size,lock_info,100]
		puts "Disk size 1 - #{disk_size}"
		db.update("UPDATE NuixReportingInfo SET CollectionStatus = ?, CaseSizeOnDisk = ?, LockInfo = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", processing_data)
		session_data = [case_guid,'','',case_name,'Get New Data','0','','','',casefolder,'',case_version,'','0',disk_size,'0',is_compound,'','',case_investigator,'','','',broker_memory,worker_count,worker_memory,evidence_name,evidence_locations,'','','',creation_date,modified_date,'','','','','','','','','','','','','','','','','','','','','','','','','','']

		# get the version of Nuix that is currently being used
		current_nuixversion = NUIX_VERSION
		# set the value for the console array
		consolevaluearray = []
		# loop over the versionmapping maps to determine if the case can be open in the version of Nuix being used
		versionmapping.each do | caseversionmap |
			# if the version map is the case version
			if caseversionmap["CaseVersion"] == case_version
				# loop through each case version and compare against the console version
				caseversionmap.each do |consoleversions|
					consoleversionval = consoleversions[0]
					if consoleversionval == "ConsoleVersions"
						consolevaluearray = consoleversions[1]
					end
				end
			end
		end

		# start an exception handler to handle any exceptions that might occur
		begin
			puts "This case - #{case_name} - version is #{case_version}"
			puts "This Console version is #{NUIX_VERSION}"
			# if the console array contains the current Nuix Version then the console can open the Nuix case
			if consolevaluearray.include? NUIX_VERSION
				puts "Console can open this case - #{case_name}"
				# open the Nuix case and pass the value of upgradecases to the open method
				nuix_case = $utilities.getCaseFactory.open("#{casefolder}",{"migrate"=>upgradecases})
				# get the getGuid from the case
				case_guid = nuix_case.getGuid
				# remove the - from the guid
				case_guid = case_guid.tr("-","")
				puts "Case guid - #{case_guid}"
				# add the processed guid string to the processed guids string
				if processed_guids == ''
					processed_guids = processed_guids + "'" + case_guid + "'"
				else
					processed_guids = processed_guids + ",'" + case_guid + "'"
				end
				# get the description of the case
				case_description = nuix_case.getDescription
				# determine if the case guid is already in the database
				count = db.scalar("select count(*) from NuixReportingInfo where CaseGUID = '#{case_guid}'")
				# determine if the case has already had reporting information collected
				count_reporting_collected = db.scalar("select count(*) from NuixReportingInfo where CaseGuid = '#{case_guid}' and ReportingInfoCollected = 'False'")
				puts "Count of Case in Database is - #{count}"
				puts "Count of Case data collected in Database is - #{count_reporting_collected}"
				# if the value is not in the database or the if the reporting info has not been collected
				if count == 0 or count_reporting_collected > 0
					# if the case was not previously added to the database add all info that was collected up to this point
					if count == 0
						session_data = [case_guid,'','',case_name,'Get New Data','0','','','',casefolder,'',case_version,'',disk_size,'0','0',is_compound,'','',case_investigator,'','','',broker_memory,worker_count,worker_memory,evidence_name,evidence_locations,'','','',creation_date,modified_date,'','','','','','','','','','','','','','','','','','','','','','','','','','','False']
						db.update("INSERT INTO NuixReportingInfo(CaseGUID, ReportLoadStart, ReportLoadEnd, CaseName, CollectionStatus, PercentComplete, ReportLoadDuration, BatchLoadInfo, ItemSets, CaseLocation, BackUpLocation, CurrentCaseVersion, UpgradedCaseVersion, CaseSizeOnDisk, CaseFileSize, CaseAuditSize, IsCompound, CasesContained, ContainedInCase, Investigator, InvestigatorSessions, InvalidSessions, InvestigatorTimeSummary, BrokerMemory, WorkerCount, WorkerMemory, EvidenceProcessed, EvidenceLocation, EvidenceCustomMetadata, MimeTypes, ItemTypes, CreationDate, ModifiedDate, LoadDataStart, LoadDataEnd, LoadTime, LoadEvents, TotalLoadTime, ProcessingSpeed, Custodians, CustodianCount, SearchTerm, SearchSize, HitCount, CustodianSearchHit, TotalItemCount, ItemCounts, OriginalItems, DuplicateItems, CaseUsers, ReportLoadTime, NuixLogLocation, OldestItem, NewestItem, Languages, CustomMetadata, CaseDescription, EvidenceDescription, IrregularItems, ReportingInfoCollected) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", session_data)
					end
					puts "Case does not exist in database or has not had reporting data collected"
					# get the name of the case
					case_name = nuix_case.getName().strip
					# create a new array to store all the case user in
					all_case_users = Array.new
					# get all the user of this case
					case_users = nuix_case.getAllUsers()
					# loop through all users
					case_users.each do |case_user|
						#add the case_user to the all_case_users array if the value is not equal to the ucrtreportingusers
						if case_user != ucrtreportinguser
							all_case_users << case_user
						end
					end
				
					# set the value of case_opened to 1 (true)
					case_opened = 1
					puts "Geting Case - #{case_name} - Statistics"
					# get the case statistics from the UCRT class getCaseStats method
					UCRT.getCaseStats(nuix_case, case_guid, showsizein, decimal_places, db)
					puts "Getting Case - #{case_name} - Size Info"
					# get the case Sizes from the UCRT class getCaseSize method
					UCRT.getCaseSize(nuix_case, case_guid, showsizein, decimal_places, db)
					puts "Getting Case - #{case_name} - Types"
					# get the case types from the UCRT class getCaseTypes method
					UCRT.getCaseTypes(nuix_case, case_guid, showsizein, decimal_places, db)
					# if the ignorecasehistory is false then get the case history
					if ignorecasehistory == "false"
						puts "Getting Case - #{case_name} - History"
						UCRT.getCaseHistory(nuix_case, case_guid, showsizein, decimal_places, includeannotations, db)
					end
					# get the itemset information
					puts "Getting - #{case_name} - Itemset Info"
					UCRT.getItemSetInfo(nuix_case, case_guid, showsizein, decimal_places, begindate, enddate, db)
					# if the include datarange is true then get all the case stats for every day from the 
					# oldest item date to the newest item date
					if includedaterange == "true"
						puts "Getting - #{case_name} - date range data"
						UCRT.getDateRangeInfo(nuix_case, case_guid, showsizein, decimal_places, db)
					end
					# get the irregular items
					puts "Getting - #{case_name} - Irregular Items Info"
					UCRT.getIrregularItems(nuix_case, case_guid, showsizein, decimal_places, db)
					# if a search term was provided search the case for the term
					if searchterm != ""
						puts "Searching - #{case_name} - for #{searchterm}"
						UCRT.getSearchCriteria(nuix_case, case_guid, searchterm, showsizein, decimal_places, db)
						if nuixexportsearchresults == "true"
							UCRT.nuixExportSearchTermResults(nuix_case, case_guid, exportfolder, searchterm, nuixeporttype, showsizein, decimal_places, db)
						end
					end
					# if a search term file was provided search the case for each value in the file
					if searchtermfile != ""
						puts "Searching - #{case_name} - for queries in #{searchtermfile}"
						UCRT.getSearchFileCriteria(nuix_case, case_guid, searchtermfile, showsizein, decimal_places, db)
						if nuixexportsearchresults == "true"
							UCRT.nuixExportSearchFileResults(nuix_case, case_guid, exportfolder, searchtermfile, nuixexporttype, showsizein, decimal_places, db)
						end
					end
					# start an exception handler to catch any exceptions that might occur
					begin
						# update the database with all values captured up to this point
						caseloadend = DateTime.now
						caseloadendtime = caseloadend.strftime("%m/%d/%Y %H:%M")
						elapsed_seconds = ((caseloadend - caseloadstart) * 24 * 60 * 60).to_i
						puts "Update Processing Data 1"
						processing_data = [caseloadstarttime, caseloadendtime, elapsed_seconds, "File System and Case Data collected","True",100]
						db.update("UPDATE NuixReportingInfo SET ReportLoadStart = ?, ReportLoadEnd = ?, ReportLoadDuration = ?, CollectionStatus = ?, ReportingInfoCollected = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", processing_data)
						nuix_case.close
					rescue StandardError => msg
						puts "Error updating Database #{msg} for - #{case_name} - #{case_guid}"
					ensure
						nuix_case.close
					end
				# if the value is in the database
				else
					open_options = {
						"startDateAfter"=> begindate.to_s,
						"startDateBefore"=> enddate.to_s,
						"type"=> "openSession"
					}
					# get the case history
					openhistory = nuix_case.getHistory(open_options)
					case_opened = 0
					# loop over the case history 
					openhistory.each do |openhist|
						# get the history user and if it is not only opened by the ucrt user open the case
						username = openhist.getUser
						if username != ucrtreportinguser
							case_opened = 1
						end
					end
					# if the case should be opened then open it and get stat
					if case_opened == 1
						puts "Case was opened between #{begindate} and #{enddate}"
						# delete the case from the reporting info 
						db.update("Delete from NuixReportingInfo where CaseGUID = '#{case_guid}'")
						session_data = [case_guid,'','',case_name,'Get New Data','0','','','',casefolder,'',case_version,'',disk_size,'0','0',is_compound,'','',case_investigator,'','','',broker_memory,worker_count,worker_memory,evidence_name,evidence_locations,'','','',creation_date,modified_date,'','','','','','','','','','','','','','','','','','','','','','','','','','']
						# Insert to collected info into the database
						db.update("INSERT INTO NuixReportingInfo(CaseGUID, ReportLoadStart, ReportLoadEnd, CaseName, CollectionStatus, PercentComplete, ReportLoadDuration, BatchLoadInfo, ItemSets, CaseLocation, BackUpLocation, CurrentCaseVersion, UpgradedCaseVersion, CaseSizeOnDisk, CaseFileSize, CaseAuditSize, IsCompound, CasesContained, ContainedInCase, Investigator, InvestigatorSessions, InvalidSessions, InvestigatorTimeSummary, BrokerMemory, WorkerCount, WorkerMemory, EvidenceProcessed, EvidenceLocation, EvidenceCustomMetadata, MimeTypes, ItemTypes, CreationDate, ModifiedDate, LoadDataStart, LoadDataEnd, LoadTime, LoadEvents, TotalLoadTime, ProcessingSpeed, Custodians, CustodianCount, SearchTerm, SearchSize, HitCount, CustodianSearchHit, TotalItemCount, ItemCounts, OriginalItems, DuplicateItems, CaseUsers, ReportLoadTime, NuixLogLocation, OldestItem, NewestItem, Languages, CustomMetadata, CaseDescription, EvidenceDescription, IrregularItems) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", session_data)
							case_opened = 1
							puts "Geting Case - #{case_name} - Statistics for case that was opened"
							UCRT.getCaseStats(nuix_case, case_guid, showsizein, decimal_places, db)
							puts "Getting Case - #{case_name} - Size Info for case that was opened"
							UCRT.getCaseSize(nuix_case, case_guid, showsizein, decimal_places, db)
							puts "Getting Case - #{case_name} - Types for case that was opened"
							UCRT.getCaseTypes(nuix_case, case_guid, showsizein, decimal_places, db)
							if ignorecasehistory == "false"
								puts "Getting Case - #{case_name} - History for case that was opened"
								UCRT.getCaseHistory(nuix_case, case_guid, showsizein, decimal_places, includeannotations, db)
							end
							puts "Getting - #{case_name} - Itemset Info for case that was opened"
							UCRT.getItemSetInfo(nuix_case, case_guid, showsizein, decimal_places, begindate, enddate, db)
							if includedaterange == "true"
								puts "Getting - #{case_name} - date range data for case that was opened"
								UCRT.getDateRangeInfo(nuix_case, case_guid, showsizein, decimal_places, db)
							end
							puts "Getting - #{case_name} - Irregular Items Info for case that was opened"
							UCRT.getIrregularItems(nuix_case, case_guid, showsizein, decimal_places, db)
							if searchterm != ""
								puts "Searching - #{case_name} - for #{searchterm}"
								UCRT.getSearchCriteria(nuix_case, case_guid, searchterm, showsizein, decimal_places, db)
								if nuixexportsearchresults == "true"
									UCRT.nuixExportSearchTermResults(nuix_case, case_guid, exportfolder, searchterm, showsizein, decimal_places, db)
								end
							end
							if searchtermfile != ""
								puts "Searching - #{case_name} - for queries in #{searchtermfile}"
								UCRT.getSearchFileCriteria(nuix_case, case_guid, searchtermfile, showsizein, decimal_places, db)
								if nuixexportsearchresults == "true"
									UCRT.nuixExportSearchFileResults(nuix_case, case_guid, exportfolder, searchtermfile, showsizein, decimal_places, db)
								end
							end
							begin
								caseloadend = DateTime.now
								caseloadendtime = caseloadend.strftime("%m/%d/%Y %H:%M")
								elapsed_seconds = ((caseloadend - caseloadstart) * 24 * 60 * 60).to_i
								processing_data = [caseloadstarttime, caseloadendtime, elapsed_seconds, "File System and Case Data collected","True",100]
								db.update("UPDATE NuixReportingInfo SET ReportLoadStart = ?, ReportLoadEnd = ?, ReportLoadDuration = ?, CollectionStatus = ?, ReportingInfoCollected = ?, PercentComplete = ? WHERE CaseGUID = '#{case_guid}'", processing_data)
								nuix_case.close
							rescue StandardError => msg
								puts "Error updating Database #{msg} - #{case_name} - #{case_guid} for case that was opened"
							ensure
								nuix_case.close
							end
						else
							puts "Case - #{case_name} was not opened between #{begindate.to_s} and #{enddate.to_s}"
							nuix_case.close
						end
					end
				else
					puts "Console cannot open - #{case_name} - #{case_guid}"
				end
			rescue StandardError => msg
				puts "Case - #{case_name} - could not be opened #{msg}"
				session_data = [case_name, case_guid, casefolder, case_version,is_compound, creation_date, casemodified, disk_size,case_investigator, worker_count, worker_memory, broker_memory, evidence_name,evidence_locations, evidence_description]
				db.update("INSERT INTO NuixReportingInfo(CaseName,CaseGUID,CaseLocation,CurrentCaseVersion,IsCompound,CreationDate,ModifiedDate,CaseSizeOnDisk,Investigator,WorkerCount,WorkerMemory,BrokerMemory,EvidenceProcessed,EvidenceLocation,EvidenceDescription) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", session_data)
			ensure

			end
		end
	end
# if the export type is cvs
if exporttype == 'csv'
	#if the export file name is not present create an export file name
	if exportfilename == ''
		timestring=DateTime.now
		timestring=timestring.strftime("%Y-%m-%d-%H-%M-%S")
		csvfilename = "#{exportdirectory}\\UCRTExportReport-#{timestring}.csv"
	# else create the export file with the name given
	else
		csvfilename = "#{exportdirectory}\\#{exportfilename}"
	end
	# if the export fields are "all"
	if csvexportfields == 'all'
		# if the exportcaseinfo is "all" then export the entire database
		if exportcaseinfo == 'all'
			csvexportstatement = "select CaseGuid,ReportLoadStart,ReportLoadEnd,CaseName,CollectionStatus,LockInfo,PercentComplete,ReportLoadDuration,BatchLoadInfo,ItemSets,CaseLocation,BackUpLocation,CurrentCaseVersion,UpgradedCaseVersion,CaseSizeOnDisk,CaseFileSize,CaseAuditSize,IsCompound,CasesContained,Investigator,InvestigatorSessions,InvalidSessions,InvestigatorTimeSummary,BrokerMemory,WorkerCount,WorkerMemory,EvidenceProcessed,EvidenceLocation,EvidenceCustomMetadata,MimeTypes,ItemTypes,CreationDate,ModifiedDate,LoadDataStart,LoadDataEnd,LoadTime,LoadEvents,TotalLoadTime,ProcessingSpeed,Custodians,CustodianCount,SearchTerm,SearchSize,HitCount,CustodianSearchHit,TotalItemCount,ItemCounts,OriginalItems,DuplicateItems,CaseUsers,ReportLoadTime,NuixLogLocation,OldestItem,NewestItem,Languages,CustomMetadata,CaseDescription,EvidenceDescription,IrregularItems from NuixReportingInfo"
		# if the exportcaseinfo is "processed" meaning just the cases that were processed in this run then export only those cases
		elsif exportcaseinfo == 'processed'
 			csvexportstatement = "select CaseGuid,ReportLoadStart,ReportLoadEnd,CaseName,CollectionStatus,LockInfo,PercentComplete,ReportLoadDuration,BatchLoadInfo,ItemSets,CaseLocation,BackUpLocation,CurrentCaseVersion,UpgradedCaseVersion,CaseSizeOnDisk,CaseFileSize,CaseAuditSize,IsCompound,CasesContained,Investigator,InvestigatorSessions,InvalidSessions,InvestigatorTimeSummary,BrokerMemory,WorkerCount,WorkerMemory,EvidenceProcessed,EvidenceLocation,EvidenceCustomMetadata,MimeTypes,ItemTypes,CreationDate,ModifiedDate,LoadDataStart,LoadDataEnd,LoadTime,LoadEvents,TotalLoadTime,ProcessingSpeed,Custodians,CustodianCount,SearchTerm,SearchSize,HitCount,CustodianSearchHit,TotalItemCount,ItemCounts,OriginalItems,DuplicateItems,CaseUsers,ReportLoadTime,NuixLogLocation,OldestItem,NewestItem,Languages,CustomMetadata,CaseDescription,EvidenceDescription,IrregularItems from NuixReportingInfo where CaseGuid in (#{processed_guids})"
		else
			csvexportstatement = "select CaseGuid,ReportLoadStart,ReportLoadEnd,CaseName,CollectionStatus,LockInfo,PercentComplete,ReportLoadDuration,BatchLoadInfo,ItemSets,CaseLocation,BackUpLocation,CurrentCaseVersion,UpgradedCaseVersion,CaseSizeOnDisk,CaseFileSize,CaseAuditSize,IsCompound,CasesContained,Investigator,InvestigatorSessions,InvalidSessions,InvestigatorTimeSummary,BrokerMemory,WorkerCount,WorkerMemory,EvidenceProcessed,EvidenceLocation,EvidenceCustomMetadata,MimeTypes,ItemTypes,CreationDate,ModifiedDate,LoadDataStart,LoadDataEnd,LoadTime,LoadEvents,TotalLoadTime,ProcessingSpeed,Custodians,CustodianCount,SearchTerm,SearchSize,HitCount,CustodianSearchHit,TotalItemCount,ItemCounts,OriginalItems,DuplicateItems,CaseUsers,ReportLoadTime,NuixLogLocation,OldestItem,NewestItem,Languages,CustomMetadata,CaseDescription,EvidenceDescription,IrregularItems from NuixReportingInfo where CaseGuid in (#{exportcaseinfo})"
		end
		puts "CVS Export Statement = #{csvexportstatement}"
		# export the values to the csv file
		CSV.open("#{csvfilename}", "wb") do |csv|
			headers_written = false
			db.query("#{csvexportstatement}") do |row|
				if !headers_written
					csv << row.keys
					headers_written = true
				end
				csv << row.values
			end
		end
	# if fields were added to the json then only export those fields
	else
		if exportcaseinfo == 'all'
			csvexportstatement = "select #{csvexportfields} from NuixReportingInfo where IsCompound = 'false'"
		elsif exportcaseinfo == 'processed'
			csvexportstatement = "select #{csvexportfields} from NuixReportingInfo where CaseGuid in (#{processed_guids}) and IsCompound = 'false'"
		else
			csvexportstatement = "select #{csvexportfields} from NuixReportingInfo where CaseGuid in (#{exportcaseinfo}) and IsCompound = 'false'"
		end
		CSV.open("#{csvfilename}", "wb") do |csv|
			headers_written = false
			db.query("#{csvexportstatement}") do |row|
				if !headers_written
					csv << row.keys
					headers_written = true
				end
				csv << row.values
			end
		end
	end
	
	# if the settings.json specify to clean up the files
	if cleanupfiles == 'true'

		puts "Cleaning up files in #{cleanup_directories}"
		cleanupdirectories = cleanup_directories.split(",")
		# loop over every directory specified in the json value
		cleanupdirectories.each do |cleanupdirectory|
			# change directory to the specific directory
			Dir.chdir(cleanupdirectory)
			# split the values of the files types to clean up
			cleanupfiletypes = cleanup_filetypes.split(",")
			# loop over each file type to clean up
			cleanupfiletypes.each do |cleanup_filetype|
				
				cleanupfileext = "**/#{cleanup_filetype}"
				puts "Cleanup File - #{cleanupfileext}"
				# get each file specified in the clean up type loop
				Dir.glob("**/#{cleanup_filetype}").each { 
					|filename|
						fileage = file_age(filename)
						puts "Filename - #{filename} - Fileage - #{fileage}" 
						# delete each file that is older then the cleanup range
						File.delete(filename) if file_age(filename) > cleanup_filerange 
				}
			end
			Dir['**/'].reverse_each { 
				|d|
				puts "Directory - #{d}"
				Dir.rmdir d if Dir.entries(d).size == 2
			}
		end
	end
end
# Get the current time so the we can report on elapsed time for the UCRT to be completed
reportloadend = DateTime.now
# Convert end time to the necessary format
reportloadendtime = reportloadend.strftime("%m/%d/%Y %H:%M")
# get the elapsed seconds from the start time and end time
elapsed_seconds = ((reportloadend - reportloadstart) * 24 * 60 * 60).to_i
# Write the elapsed time to the log file
puts "Total Report Time #{elapsed_seconds}"
