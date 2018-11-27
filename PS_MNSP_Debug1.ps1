
################################################################################################################################################################
# Quest - Migrator for Notes to SharePoint
################################################################################################################################################################
Import-Module "C:\Program Files (x86)\Quest\Migrator for Notes to SharePoint\Bin\Quest.NSP.SharePoint.Common.dll"
Import-Module "C:\Program Files (x86)\Quest\Migrator for Notes to SharePoint\Bin\Quest.NSP.Transform.Common.dll"
Import-Module "C:\Program Files (x86)\Quest\Migrator for Notes to SharePoint\Bin\Quest.NSP.Migrator.Common.dll"
################################################################################################################################################################

<#

NotesColumnType
	Item,
	ViewColumn,
	Formula,
	Document,
	Unid,
	ParentItem,
	ParentFormula,
	ParentUnid,
	NoteLink,
	RichText,
	Render,
	Attachment,
	AttachmentName,
	AttachmentId,
	AttachmentInfo,
	AttachmentBlocked,
	AttachmentLinks,
	Image,
	ImageType,
	OleObject,
	OleObjectClass,
	OleObjectType

NotesColumnOption
	None,
	Multi,
	Flat,
	Xml,
	XmlNoBinary,
	Html,
	Cd,
	Mime

ColumnDataType
	None,
	String,
	HtmlString,
	Number,
	Date,
	Binary,
	User


#>



[Quest.NSP.Transform.QueryFactory] $qryFactory = [Quest.NSP.Transform.QueryFactory]::LoadFactory([Quest.NSP.Transform.DataSourceType]::Notes)
[Quest.NSP.Migrator.Options] $opts = [Quest.NSP.Migrator.Options]::Load()

[Quest.NSP.SharePoint.TransferJob] $transJob = New-Object Quest.NSP.SharePoint.TransferJob($qryFactory, [Quest.NSP.SharePoint.TargetType]::SharePoint)
$transJob.JobOptions.PreserveDates = $true

$opts.UserOptions.InitializeJobWithUserDefaults($transJob)


[Quest.NSP.Transform.QuerySource] $qrySrc = $qryFactory.CreateQuerySource("DANLDA01/DA-DE/Server/Degussa","Rohmax/Supply Chain/RXGSCT.nsf")

[Quest.NSP.Transform.NotesGlobalQueryOptions] $nGlbQryOpts = New-Object Quest.NSP.Transform.NotesGlobalQueryOptions

$qrySrc.MergeOptions($nGlbQryOpts)

