#!/usr/bin/perl 
use strict;  
use warnings;   

use FileHandle;
use DirHandle;

use Archive::Zip qw( :ERROR_CODES :CONSTANTS );
use Archive::Zip::MemberRead;
use Data::Dumper;
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;
use XML::LibXML;
use Text::CSV;
use DateTime;



#  Make sure an argument is passed to the script.  This will be the 
#  map size to find zip files for and process into a spreadsheet;
my ($mapsize) = @ARGV;
 if ( !defined($mapsize) || !($mapsize =~/4096|8192/))
{
	die "spreadsheet.pl <size>\n";
}

my %FORMATS;


my @top7;
my @top15;

my %top7;
my %top15;

my @skyscrapers;
my %skyscrapers;

my @traders;
my %traders;

my @tier4;
my %tier4;

my @tier5;
my %tier5;

my @stores;
my %stores;

my @specials;
my %specials;

my %POITransations;


my $F7D2D = $ENV{'F7D2D'};

my $rootDir = '/home/jfraz/seedGen';
my $specialsDir = $rootDir . "/special";

my $previewsDir = $F7D2D . "/previews";
my $excelFile = $previewsDir . "/SeedComparison-$mapsize.xlsx";


my $previewDirHandle = new DirHandle $previewsDir;

my $workbook= Excel::Writer::XLSX->new( $excelFile ) || die "Unable to create sheet.  Is the worksheet open?";  
my $worksheet;

my $fileCounter = 1;
my $worldSeedsCount = 0;

my $zipLimit = 1400;

my %allPrefabsDetails;

sub load_specials()
{
	my $currentFile = $specialsDir . "/top7.txt";
	@top7 = load_special_file($currentFile);
	# now, create a hash for quick lookup using exists
	%top7 = map { $_ => 1 } @top7;
	
	$currentFile = $specialsDir . "/top15.txt";	
	@top15 = load_special_file($currentFile);
	%top15 = map { $_ => 1 } @top15;
	
	$currentFile = $specialsDir . "/skyscrapers.txt";	
	@skyscrapers = load_special_file($currentFile);
	%skyscrapers = map { $_ => 1 } @skyscrapers;

	$currentFile = $specialsDir . "/tier4.txt";	
	@tier4 = load_special_file($currentFile);
	%tier4 = map { $_ => 1 } @tier4;	

	$currentFile = $specialsDir . "/tier5.txt";	
	@tier5 = load_special_file($currentFile);
	%tier5 = map { $_ => 1 } @tier5;		
	
	$currentFile = $specialsDir . "/stores.txt";	
	@stores = load_special_file($currentFile);
	%stores = map { $_ => 1 } @stores;		
	
	$currentFile = $specialsDir . "/special.txt";	
	@specials = load_special_file($currentFile);
	%specials = map { $_ => 1 } @stores;			
	
}


sub load_translations()
{
	my $csv = Text::CSV->new({ sep_char => ',' });

	while (my $line = <DATA>) 
	{
	  chomp $line;
	 
	  if ($csv->parse($line)) 
	  {
		my ($key, $text) = $csv->fields();
		$POITransations{$key} = { text => $text, tier => "unknown"};
	  }
	}
}


# This is a generic loader for each specials file
# it takes a fully qualified filename path as an input
# and returns a list of all records within that file
sub load_special_file()
{
	# pull the passed in filename arg and define our return list variable
	my $specialFile = shift @_; 
	my @returnList =();
	
	my $specialsFileFH = FileHandle->new($specialFile, "r"); 
	if (defined $specialsFileFH)
	{
		# read each line from the filehandle and put it into the returnlist 
		while (my $line = $specialsFileFH->getline())
		{
			# remove the newline character from the line
			chomp($line);
			push(@returnList, $line);
		}
		return @returnList;
	}	
}

sub loalPrefabDetails()
{
	my $dataFolder = $F7D2D . "/Data";
	
	my $dom = XML::LibXML->load_xml(location => $dataFolder . "/Config/rwgmixer.xml");	
	my (@allPrefabsList) = $dom->findnodes('//prefab_rule[@name="prefabList"]/prefab');
	foreach my $prefabNode (@allPrefabsList)
	{
		my $name = $prefabNode->getAttribute('name');	
		
		my $prefabDom = XML::LibXML->load_xml(location => $dataFolder . "/Prefabs/$name.xml");	
		
		my $questTiers = $prefabDom->findvalue('//prefab/property[@name ="DifficultyTier"]/@value');
		my $questTags = $prefabDom->findvalue('//prefab/property[@name ="QuestTags"]/@value');
		
		$allPrefabsDetails{$name}{"QuestTier"} = $questTiers;
		$allPrefabsDetails{$name}{"isFetch"} = ($questTags =~ /fetch/) ? 1 : 0;
		$allPrefabsDetails{$name}{"isClear"} = ($questTags =~ /clear/) ? 1 : 0;
		$allPrefabsDetails{$name}{"isHiddenCache"} = ($questTags =~ /hidden_cache/) ? 1 : 0;
		
		#print Dumper(\%allPrefabsDetails);
		
		#printf "%-35s%-3s\t%-35s\n", $name, $questTiers, $questTags; 		
		
		
		
	}
	
}


sub buildFirstSheet()
{
	my $worksheet = $workbook->add_worksheet("8K Map Seeds Info");  
	my @headers = ("Seed", "Map Name", "Prefabs", "Unique\nPOIs"
	, "Tier4", "Tier5", "Top7", "Top15", "Stores", "Tower"
	, "POI %\nForest"
	, "POI %\nSnow"
	, "POI %\nDesert"
	, "POI %\nBurnt"
	, "POI %\nWaste"
	, ""
	, "Land %\nForest"
	, "Land %\nSnow"
	, "Land %\nDesert"
	, "Land %\nBurnt"
	, "Land %\nWaste"
	, "", "Missing Top 15", "Gen Date"

	);
	
	my $format = $workbook->add_format(
    border => 6,
    valign => 'vcenter',
    align  => 'center',
	);
 
	$worksheet->write('A1', \@headers, $FORMATS{'HEADER'} );
	$worksheet->write('P1', "", $FORMATS{'COLSEP'} );
	$worksheet->write('V1', "", $FORMATS{'COLSEP'} );
	
	
	$worksheet->set_column( "A1:B1", 27 );
	$worksheet->set_column( "C1:J1", 8 );
	$worksheet->set_column( "K1:O1", 7 );
	$worksheet->set_column( "Q1:U1", 7 );
	$worksheet->set_column( "W1:W1", 35 );
	$worksheet->set_column( "X1:X1", 14 );
	
	$worksheet->set_column( "V1:V1", 4, $FORMATS{'COLSEP'} );
	$worksheet->set_column( "P1:P1", 4, $FORMATS{'COLSEP'} );
	
	$worksheet->freeze_panes( 1, 1 ); 
	$worksheet->set_row(0, 30); 
	
	$worksheet->conditional_formatting( 'C:C',
    {
        type      => '3_color_scale', 
        min_color => "#FF0000",
		mid_color => "#FFFFFF",		
        max_color => "#85CD00"
    } );	
			

	$worksheet->conditional_formatting( 'D:D',
    {
        type      => '3_color_scale', 
        min_color => "#FF0000",
		mid_color => "#FFFFFF",		
        max_color => "#85CD00"
    } );	
		
	$worksheet->conditional_formatting( 'E:E',
    {
        type      => '3_color_scale', 
        min_color => "#FF0000",
		mid_color => "#FFFFFF",		
        max_color => "#85CD00"
    } );
	$worksheet->conditional_formatting( 'F:F',
    {
        type      => '3_color_scale', 
        min_color => "#FF0000",
		mid_color => "#FFFFFF",		
        max_color => "#85CD00"
    } );	

	
	$worksheet->conditional_formatting( 'D:D',
    {
        type      => '3_color_scale', 
        min_color => "#FF0000",
		mid_color => "#FFFFFF",		
        max_color => "#85CD00"
    } );
	$worksheet->conditional_formatting( 'G:G',
    {
        type      => '3_color_scale',
        min_color => "#FF0000",
		mid_color => "#FFFFFF",		
        max_color => "#85CD00"
    });
 	$worksheet->conditional_formatting( 'H:H',
    {
        type      => '3_color_scale', 
        min_color => "#FF0000",
		mid_color => "#FFFFFF",		
        max_color => "#85CD00"
    } );
		
	$worksheet->conditional_formatting( 'I:I',
    {
        type      => '3_color_scale', 
        min_color => "#FF0000",
		mid_color => "#FFFFFF",		
        max_color => "#85CD00"
    } );
	$worksheet->conditional_formatting( 'J:J',
    {
        type      => '3_color_scale', 
        min_color => "#FF0000",
		mid_color => "#FFFFFF",		
        max_color => "#85CD00"
    } );

	
	return $worksheet;
}


sub buildHeaderFormat()
{
	my $format = $workbook->add_format(
		border => 1,
    	valign => 'vcenter');
	$format->set_bold();
	$format->set_align( 'center' );
	return $format;
}

sub writeBaseSheetRow()
{
	my ($worldSeedCount, $seed, $mapname, $gendate, $prefabs
	, $uniques, $tier4, $tier5, $top7, $top15
	, $stores, $skyscrapers) = @_;

	$worksheet->write($worldSeedCount, 0, $seed,$FORMATS{'DATA'});
	$worksheet->write($worldSeedCount, 1, $mapname,$FORMATS{'DATA'});
	
	
	
	$worksheet->write($worldSeedCount, 2, $prefabs,$FORMATS{'DATA'});
	$worksheet->write($worldSeedCount, 3, $uniques,$FORMATS{'DATA'});
	
	$worksheet->write($worldSeedCount, 4, $tier4,$FORMATS{'DATA'});
	$worksheet->write($worldSeedCount, 5, $tier5,$FORMATS{'DATA'});
	
	
	$worksheet->write($worldSeedCount, 6, $top7,$FORMATS{'DATA'});	
	$worksheet->write($worldSeedCount, 7, $top15,$FORMATS{'DATA'});
	
	$worksheet->write($worldSeedCount, 8, $stores,$FORMATS{'DATA'});	
	$worksheet->write($worldSeedCount, 9, $skyscrapers,$FORMATS{'DATA'});
	
	# the gen date
	$worksheet->write($worldSeedCount, 25, $gendate,$FORMATS{'DATA'});
}



sub processZipFiles()
{
	if (defined $previewDirHandle)
	{
		while (defined(my $currentZipFile  = $previewDirHandle->read)) 
		{
			next if ($fileCounter >= $zipLimit);	
			my $worldSeedName;
			
			if (! -d $currentZipFile )
			{
				if($currentZipFile =~ /\.zip$/)
				{
					processZipFile($currentZipFile);
					$fileCounter++;
				}
			}
			
			undef $worldSeedName;
			
		}
	}
}

sub processPOICommentCounts($$$)
{
	# We have a scalar, an array
	# , and a hash passed by reference
	my $columnNumber = shift;
	my $compareList = shift;
	my $seedPrefabs = shift;
	
	my $pois ="";
	
	foreach my $items (@$compareList)
	{
		my $poiName = $POITransations{$items}{"text"} || $items;
		if(exists($seedPrefabs->{$items}))
		{
			$pois .= sprintf "%s\t%-40s\n", $seedPrefabs->{$items}, $poiName;
		}
		else 
		{
			$pois .= sprintf "%s\t%-40s\n", "X" , $poiName;
		}
		$worksheet->write_comment($fileCounter, $columnNumber, $pois
			, height 	=> 250
			, width 	=> 375
			, font_size => 13);
	}
}

sub processZipFile()
{
	my $zipFileName = shift;
	#	print "File: " . $zipFileName . "\n";
	my ($basefilename) = ($zipFileName =~ /(.+?)\.zip/);
	my @members; 

	my $zipFile  = Archive::Zip->new();
	
	unless ( $zipFile->read( $previewsDir ."/$zipFileName" ) == AZ_OK ) {
    die 'read error'; }

	my ($device, $inode, $mode, $nlink, $uid
		, $gid, $rdev, $size, $atime, $mtime
		, $ctime, $blksize, $blocks) 
		= stat($previewsDir ."/$zipFileName" );

	my $dt = DateTime->from_epoch( epoch => $mtime );
	@members = $zipFile->memberNames();
 
 	my ($top7Count, $top15Count, %seedPrefabs) = processPrefabs($zipFile, $basefilename, $dt);	
	my %biomeData = processBiomeFile($zipFile, $basefilename);
	#print Dumper(\%POITransations);
	
	if(($top7Count >=7 && $top15Count >=10) || ($top7Count >=6 && $top15Count >=11))
	{
		processSeedDescription($basefilename, %seedPrefabs);
		#processSeedSheet($basefilename, %seedPrefabs);
		my $missingPrefabs;
		my $temp;
		my $hasMissing =0;
		
		my $pois ="";
		
		processPOICommentCounts(4 , \@tier4, \%seedPrefabs );
		processPOICommentCounts(5 , \@tier5, \%seedPrefabs );
		processPOICommentCounts(6 , \@top7, \%seedPrefabs );
		processPOICommentCounts(7 , \@top15, \%seedPrefabs );
		processPOICommentCounts(8 , \@stores, \%seedPrefabs );
		processPOICommentCounts(9 , \@skyscrapers, \%seedPrefabs );
		#foreach my $items (@top7)
		#{
			#unless (exists ($seedPrefabs{$items}))
			#{
			#	$hasMissing=1;
			#	#print $items . " - " . $POITransations{$items} . "\n";
			#	my $missing = $POITransations{$items}{"text"} || $items;
			#	$temp .= "\t$missing\n";
			#}
		#	my $poiName = $POITransations{$items}{"text"} || $items;		
		#	if(exists($seedPrefabs{$items}))
		#	{
		#		$pois .= sprintf "%s\t%-40s\n", $seedPrefabs{$items}, $poiName;
		#	}
		#	else 
		#	{
		#		$pois .= sprintf "%s\t%-40s\n", "X" , $poiName;
		#	}
			#$worksheet->write_comment($fileCounter, 6, $pois
			#	, height 	=> 250
			#	, width 	=> 375
			#	, font_size => 13);
			
			
		#}
		#undef $pois ;
		
		if($hasMissing ==1)
		{
			$missingPrefabs .= "Top7: \n";
			$missingPrefabs .= $temp;
			$hasMissing =0;
			$temp = "";
		}
		
		foreach my $items (@top15)
		{
			unless (exists ($seedPrefabs{$items}))
			{
				$hasMissing=1;
				#print $items . " - " . $POITransations{$items} . "\n";
				my $missing = $POITransations{$items}{"text"} || $items;
				$temp  .= "\t$missing\n";
			}
		}
		if($hasMissing ==1)
		{
			$missingPrefabs .= "Top15: \n";
			$missingPrefabs .= $temp;
			$hasMissing =0;
			$temp = "";
		}
		

		foreach my $items (@tier4)
		{
			unless (exists ($seedPrefabs{$items}))
			{
				$hasMissing=1;
				#print $items . " - " . $POITransations{$items} . "\n";
				my $missing = $POITransations{$items}{"text"} || $items;
				$temp  .= "\t$missing\n";
			}
		}
		if($hasMissing ==1)
		{
			$missingPrefabs .= "Tier4: \n";
			$missingPrefabs .= $temp;
			$hasMissing =0;
			$temp = "";
		}
 
		foreach my $items (@tier5)
		{
			unless (exists ($seedPrefabs{$items}))
			{
				$hasMissing=1;
				#print $items . " - " . $POITransations{$items} . "\n";
				my $missing = $POITransations{$items}{"text"} || $items;
				$temp  .= "\t$missing\n";
			}
		}
		if($hasMissing ==1)
		{
			$missingPrefabs .= "Tier5: \n";
			$missingPrefabs .= $temp;
			$hasMissing =0;
			$temp = "";
		}

		chomp $missingPrefabs;
		$worksheet->write($fileCounter, 22, "see comment", $FORMATS{'DATA'});
		$worksheet->write_comment($fileCounter, 22, $missingPrefabs
				, height 	=> 250
				, width 	=> 375
				, font_size => 13);
	}
	else
	{
		$worksheet->write($fileCounter, 22, "" ,$FORMATS{'DATA'});	
	}
	
	my $countyName = processCountyName($zipFile, $basefilename);
	$worksheet->write($fileCounter, 1, "$countyName",$FORMATS{'DATA'});
	
	
	
	my $totalPOIs = $biomeData{"POICount"} || 0;
	my $totalArea = $biomeData{"TotalArea"} || 0;
	#my $forestPOIs, $snowPOIS;
	if($totalPOIs > 0 && $totalArea >0)
	{
		

		
		my $forestPOIs 		= sprintf "%.1f", (($biomeData{"Forest"}{"Total"}/$totalPOIs)*100); 
		my $forestLand 		= sprintf "%.1f", (($biomeData{"Forest"}{"LandArea"}/$totalArea)*100);
		
		my $snowPOIs 		= sprintf "%.1f", (($biomeData{"Snow"}{"Total"}/$totalPOIs)*100);
		my $snowLand 		= sprintf "%.1f", (($biomeData{"Snow"}{"LandArea"}/$totalArea)*100);
	
		my $desertPOIs 		= sprintf "%.1f", (($biomeData{"Desert"}{"Total"}/$totalPOIs)*100);
		my $desertLand 		= sprintf "%.1f", (($biomeData{"Desert"}{"LandArea"}/$totalArea)*100);

		my $burntPOIs 		= sprintf "%.1f", (($biomeData{"Burnt Forest"}{"Total"}/$totalPOIs)*100);
		my $burntLand 		= sprintf "%.1f", (($biomeData{"Burnt Forest"}{"LandArea"}/$totalArea)*100);
		
		my $wastelandPOIs 	= sprintf "%.1f", (($biomeData{"Wasteland"}{"Total"}/$totalPOIs)*100);
		my $wastelandLand 	= sprintf "%.1f", (($biomeData{"Wasteland"}{"LandArea"}/$totalArea)*100);
				
		$worksheet->write($fileCounter, 10, "$forestPOIs",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 11, "$snowPOIs",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 12, "$desertPOIs",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 13, "$burntPOIs",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 14, "$wastelandPOIs",$FORMATS{'DATA'});
		
		$worksheet->write($fileCounter, 15, "", $FORMATS{'COLSEP'});
		
		$worksheet->write($fileCounter, 16, "$forestLand",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 17, "$snowLand",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 18, "$desertLand",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 19, "$burntLand",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 20, "$wastelandLand",$FORMATS{'DATA'});
	}	
	else 
	{
		$worksheet->write($fileCounter, 10, "",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 11, "",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 12, "",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 13, "",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 14, "",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 15, "",$FORMATS{'COLSEP'});
		$worksheet->write($fileCounter, 16, "",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 17, "",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 18, "",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 19, "",$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 20, "",$FORMATS{'DATA'});
	}
}

sub processSeedSheet()
{
	my $basefilename = shift;
	my %POIS = @_;
	my ($seedName) = $basefilename =~ /(.+?)\-\d+/;
	my $sheetName = substr($seedName, 0, 30);
	my $sheet = $workbook->add_worksheet($sheetName);
	#$worksheet ->write()

	$worksheet->write_url($fileCounter, 0,   "internal:$sheetName!A1", $FORMATS{'SHEETLINK'}
	, $seedName );
	#$worksheet->write($fileCounter, 0, $seedName,$FORMATS{'DATA'});
 
 
}

sub processSeedDescription()
{
	my $basefilename = shift;
	my %POIS = @_;
	my $fileName = $previewsDir . "/$mapsize/" . $basefilename . "-Details.txt";
	my $descriptionFile  = FileHandle->new($fileName, "w");
	
	my @specialsOrder = ('skyscraper_02', 'store_book_01', 'store_book_02', 'store_gun_01', 
		'store_gun_02', 'church_01', 'bombshelter_01', 'store_hardware_01', 'store_hardware_02', 
		'store_electronics_01', 'store_electronics_02', 'store_clothing_02', 'store_autoparts_01', 
		'installation_red_mesa', 'store_grocery_02');
	my %specials =  map { $_ => 1 } @specialsOrder;
	if (defined $descriptionFile) 
	{
		foreach my $specialPOI (@specialsOrder)
		{
			if(exists($POIS{$specialPOI}))
			{
				$descriptionFile->printf("%-10s%-45s%-35s\n", $POIS{$specialPOI}, $POITransations{$specialPOI}{"text"}, $specialPOI );
			}
			else
			{
				$descriptionFile->printf("%-10s%-45s%-35s\n", 0, $POITransations{$specialPOI}{"text"}, $specialPOI );
			}
		}
		foreach my $item (sort keys(%POIS))
		{
			if(! exists($specials{$item}))
			{
				$descriptionFile->printf("%-10s%-45s%-35s\n", $POIS{$item}, $POITransations{$item}{"text"}, $item );
			}
		}
		
	}

	
}



sub processCountyName()
{
	my $zipFile = shift;
	my $basefilename = shift;
		
	my $worldNameFH  = Archive::Zip::MemberRead->new($zipFile,"$basefilename.txt");
	if (defined ($worldNameFH))
	{
		my $line = $worldNameFH->getline();
		chomp $line;
		return $line;
	}

}

sub processPrefabs()
{
	my $zipFile = shift;
	my $basefilename = shift;
	my $datestamp = shift;
	my ($seedName) = $basefilename =~ /(.+?)\-\d+/;
	
	my $prefabs  = Archive::Zip::MemberRead->new($zipFile,"$basefilename.xml");
	
	my $xmlString ="";
	my $buffer;
	my ($totalPreFabs, $uniques);
	if (defined ($prefabs))
	{
		my %prefabs; 
		my $read = $prefabs->read($buffer, 32*1024);
		
		my $dom = XML::LibXML->load_xml(string => $buffer);	
		$totalPreFabs = $dom->findvalue('count(//decoration)');
		my (@allPrefabsList) = $dom->findnodes('/prefabs/decoration');
		foreach my $decoration (@allPrefabsList)
		{
			my $name = $decoration->getAttribute('name');
			$prefabs{$name}++;
		}
		$uniques = scalar(keys(%prefabs));
			
		my %tops = (7 =>0,  15 => 0, 't4' => 0, 't5' => 0
		, 'stores' => 0, 'skyscrapers' => 0);
		foreach my $topitem (keys(%top7))
		{
			if(exists($prefabs{$topitem}))
			{
				$tops{7}++;
			}
		}
		foreach my $topitem (keys(%top15))
		{
			if(exists($prefabs{$topitem}))
			{
				$tops{15}++;
			}
		}
		foreach my $topitem (keys(%tier4))
		{
			if(exists($prefabs{$topitem}))
			{
				$tops{'t4'}++;
			}
		}		
		foreach my $topitem (keys(%tier5))
		{
			if(exists($prefabs{$topitem}))
			{
				$tops{'t5'}++;
			}
		}	
		foreach my $topitem (keys(%stores))
		{
			if(exists($prefabs{$topitem}))
			{
				$tops{'stores'}++;
			}
		}	
		foreach my $topitem (keys(%skyscrapers))
		{
			if(exists($prefabs{$topitem}))
			{
				$tops{'skyscrapers'}++;
			}
		}	
		

		
		$worksheet->write($fileCounter, 0, $seedName,$FORMATS{'DATA'});
		
			
		$worksheet->write($fileCounter, 2, $totalPreFabs,$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 3, $uniques,$FORMATS{'DATA'});		
		$worksheet->write($fileCounter, 4, $tops{'t4'},$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 5, $tops{'t5'},$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 6, $tops{7},$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 7, $tops{15},$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 8, $tops{'stores'},$FORMATS{'DATA'});
		$worksheet->write($fileCounter, 9, $tops{'skyscrapers'} ,$FORMATS{'DATA'});
		
		
		$worksheet->write_date_time($fileCounter, 23, $datestamp, $FORMATS{'CREATEDATE'} );	
		return $tops{7}, $tops{15}, %prefabs;
	}	
}



sub processBiomeFile()
{
	my $zipFile = shift;
	my $basefilename = shift;
	my %seedBiomeData = ('Snow' => {'Total', 0, 'Area', 0, 'PrefabArea' => 0}
								, 'Desert' => {'Total', 0, 'Area', 0, 'PrefabArea' => 0} 
								, 'Burnt Forest' => {'Total', 0, 'Area', 0, 'PrefabArea' => 0} 
								, 'Forest' => {'Total', 0, 'Area', 0, 'PrefabArea' => 0} 
								, 'Wasteland' => {'Total', 0, 'Area', 0, 'PrefabArea' => 0} 
								, "TotalArea" => 0
								, "POICount" => 0);
			
	if ($zipFile->membersMatching( 'biome\-count\.txt' )> 0)
	{	
		my $biomeFile  = Archive::Zip::MemberRead->new($zipFile,"$basefilename-biome-count.txt");
		if (defined ($biomeFile))
		{
			my ($sumTotal, $sumarea);
			
			my $biomeCSV = Text::CSV->new({ sep_char => ','});
			while (defined(my $line = $biomeFile->getline()))
			{
				chomp $line;
				next if ($biomeFile->input_line_number == 1);
				
				if( $biomeCSV->parse($line)) 
				{
					my ($biomeName, $totals, $prefabarea ,$landArea) = $biomeCSV->fields();
					if(defined($biomeName) and !($biomeName eq ""))
					{
						$seedBiomeData{$biomeName}{'Total'} = $totals || 0;
						$seedBiomeData{$biomeName}{'LandArea'} = $landArea || 0;
						$seedBiomeData{$biomeName}{'PrebabArea'} = $prefabarea || 0;
						$seedBiomeData{'TotalArea'}+= $landArea || 0;
						$seedBiomeData{'POICount'}+= $totals || 0;
					}
				}
			}
			#print $basefilename . "\n";
			#print Dumper(\%seedBiomeData);
			
		}
	}	
	return %seedBiomeData;
}

sub buildColSepFormat()
{
	my $format = $workbook->add_format(border => 1);
	$format->set_pattern(18);
	$format->set_bg_color( '#000000' );
	#$format->set_fg_color('#FF00FF'); 
	return $format;
}

sub buildLinkFormat()
{
	my $format = $workbook->add_format(border => 1);
	$format->set_bold();
	#$format->set_bg_color('#4169E1'); 
	#$format->set_bg_color('#4169E1'); 
	
	return $format;
}

loalPrefabDetails();
load_translations();
load_specials();



 




$FORMATS{'HEADER'} = buildHeaderFormat();
$FORMATS{'CREATEDATE'} = $workbook->add_format( num_format =>'yyyy-mm-dd', border => 1);
$FORMATS{'DATA'} = $workbook->add_format(border => 1);
$FORMATS{'COLSEP'} = buildColSepFormat();
$FORMATS{'SHEETLINK'} = buildLinkFormat();


$worksheet = buildFirstSheet();
#$worksheet->set_column( "V:V", 12 , $FORMATS{'CREATEDATE'});

processZipFiles();

#$worksheet->write('Z1', "done", $FORMATS{'COLSEP'} );
 
$workbook->close;  
 


__DATA__
bombshelter_01,Bombshelter - Barts Salvage
church_01,Church - Large multi-story
installation_red_mesa,Red Mesa Missle Silo
sawmill_01,Sawmill
school_01,School - Navezgane Highschool
school_daycare_01,School - Poopy Pants Daycare
school_k6_01,School - Favels Acadamy
skate_park_01,Skate Park
skyscraper_01,Dishong Tower
skyscraper_02,Crack a Book Tower
skyscraper_03,Higashi Tower
skyscraper_04,Joe Bros Builders
factory_lg_01,Factory - Shameway Foods
factory_lg_02,Factory - Shotgun Messiah
apartment_brick_6_flr,Apartment - Brick 6 Floor
apartment_adobe_red_5_flr,Apartment - Red Adobe 5 Floor
house_old_victorian_03,House - Old Victorian - Radio Towers
utility_waterworks_01,Utility - Waterworks
hospital_01,Hospital
store_autoparts_01,OReally Auto Parts
store_bank_01,Piggy Bank
store_book_01,Crack a Book - Large
store_book_02,Crack a Book - Small
store_clothing_01,Savage Country - Small
store_clothing_02,Savage Country - Large
store_electronics_01,Mo Power Electronics - Small
store_electronics_02,Mo Power Electronics - Large
store_grocery_02,Shamway Foods - Large
store_grocery_01,Shamway Foods - Small
store_gun_01,Shotgun Messiah - Shooting Range
store_gun_02,Shotgun Messiah - Small
store_hardware_01,Working Stiffs - Large
store_hardware_02,Working Stiffs - Small
store_laundry_01,Store - Zacharys Laundromat
store_pawn_01,Store - Vick Garrisons Pawn and Loan
store_pharmacy_01,Store - Pop n Pills - Small
store_pharmacy_02,Store - Pop n Pills - Large
store_salon,Store - Zoe Salon
trader_bob,Trader - Bob
trader_hugh,Trader - Hugh
trader_jen,Trader - Jen
trader_joel,Trader - Joel
trader_rekt,Trader - Rekt
army_barracks_01,Army Barracks - Fortified
army_camp_01,Army Camp - Large
army_camp_02,Army Camp - Medical
army_camp_03,Army Camp - Small
bar_sm_01,Buzzs Bar
bar_stripclub_01,The Boobie Trap
bar_theater_01,Brothers Theatre
blueberryfield_sm,Blueberry Field
bombshelter_lg_01,Bombshelter - Junkyard
bombshelter_md_01,Bombshelter - Well
business_old_03,Business - The Bear Den
business_old_02,Business - Pawn Shop
business_old_01,Business - Bobs Bakery
business_old_04,Business - Aldos Cabinets
business_old_05,Business - Butcher Petes
business_old_06,Business - Special Totz
business_old_07,Business - Doggos
business_old_08,Business - PB Paper Mill
business_strip_old_01,Business Strip - 
business_strip_old_02,Business Strip - Crack a Book
business_strip_old_03,Business Strip - Pop n Pills/Shamway
business_strip_old_04,Business Strip - Shotgun Messiah
carlot_01,Carlot - Joeys Carlot
carlot_02,Carlot - Carls Cars
church_graveyard1,Church - Graveyard
church_sm_01,Church - Small
courthouse_med_01,Courthouse
departure_city_blk_plaze,City Block Plaza
diner_01,Diner - Small
diner_02,Diner - Bobs Cafe
diner_03,Diner - Dels Cafe
fastfood_01,Fastfood - Prowling Petes
fastfood_02,Fastfood - Hurry Harrys
fastfood_03,Fastfood - Berserk Bills
fastfood_04,Fastfood - Fats Food
fire_station_01,Fire Station - Small
fire_station_02,Fire Station - Large Red
football_stadium,Stadium - Navezganes Coliseum
funeral_home_01,Funeral Home - Am I Gone
garage_05,Garage - Shade Tree Auto
garage_07,Garage - Construction
gas_station1,Pass-n-Gas - 1
gas_station2,Pass-n-Gas - 2
gas_station3,Pass-n-Gas - 3
gas_station4,Pass-n-Gas - 4
gas_station5,Pass-n-Gas - Large
gas_station6,Pass-n-Gas - 6
gas_station7,Pass-n-Gas - 7
gas_station8,Pass-n-Gas - Tiny
gas_station9,Pass-n-Gas - Medium
hotel_new_01,Hotel - New
hotel_ostrich,Hotel - Ostrich
hotel_roadside_01,Hotel - Motel Eight
hotel_roadside_02,Hotel - Days End Suites
house_construction_01,House - Construction
house_construction_02,House - Construction - Basement
house_modern_05,House - Spanish - Book House
house_old_gambrel_04,House - Bobs Boars/Carls Corn
house_old_bungalow_11,House - Underground Shelter
house_old_modular_02,House - Gearz
house_old_mansard_06,House - Fates Motel
house_old_tudor_01,House - Underground Caves
oldwest_business_01,Old West Business - Biskitz
oldwest_business_02,Old West Business - BJ Welding
oldwest_business_03,Old West Business - Pine Shards????
oldwest_business_04,Old West Business - Bath Haus
oldwest_business_05,Old West Business - Oil
oldwest_business_06,Old West Business - Spitz Seeds
oldwest_business_07,Old West Business - Erics Stuff
oldwest_business_08,Old West Business - Undie Takers
oldwest_business_09,Old West Business - Earls
oldwest_business_10,Old West Business - Rifle Martyrs
oldwest_business_11,Old West Business - Skamson Grocery
oldwest_business_12,Old West Business - Coles Books
oldwest_business_13,Old West Business - Swiggin Serum
oldwest_business_14,Old West Business - Lathan Hardware
oldwest_church,Old West - Church
oldwest_cole_factory,Old West - Dump n Lung Coal
oldwest_jail,Old West - Sheriff
oldwest_stables,Old West - Lazy H
oldwest_strip_01,Old West Strip - Dukes General
oldwest_strip_02,Old West Strip - Morning Lumber
oldwest_strip_03,Old West Strip - Berts Brewery
oldwest_strip_04,Old West Strip - Metal Works