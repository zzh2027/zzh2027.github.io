<!DOCTYPE html>
<html>
<head>
<title>README.md - PyAmazonCACerts - Code Browser</title>
<meta content='width=device-width, initial-scale=1' name='viewport'>
<meta content='true' name='use-sentry'>
<meta content='612b9c7b-dd7b-41c8-b9f9-86bee7acb629' name='request-id'>
<meta name="csrf-param" content="authenticity_token" />
<meta name="csrf-token" content="x7jPlaCqYVlYn2B2OCJYce2jOExM8MYD4t+kzTIF0Ek2AmPcCkoFVm1mmTqQmWXO6BGVb9RDcVPHBTxCUAn3lA==" />
<meta content='IE=edge' http-equiv='X-UA-Compatible'>
<link rel="shortcut icon" type="image/x-icon" href="https://internal-cdn.amazon.com/code.amazon.com/pub/assets/cdn/favicon-ab8a20c61652d108f6639a5452236eb5.ico" />
<link rel="stylesheet" media="all" href="//internal-cdn.amazon.com/oneg.amazon.com/assets/3.2.4/css/application.min.css" />
<link rel="stylesheet" media="all" href="https://internal-cdn.amazon.com/code.amazon.com/pub/assets/cdn/vendor-aeee840f8238d11380f6625bc01a5ff9.css" />
<link rel="stylesheet" media="all" href="https://internal-cdn.amazon.com/code.amazon.com/pub/assets/cdn/application-oneg-d87300855f7e58bb941be66f93e490fd.css" />
<link rel="stylesheet" media="all" href="https://internal-cdn.amazon.com/is-it-down.amazon.com/stylesheets/stripe.css" />
<style>
  /* line 1, (__TEMPLATE__) */
  .absolute-time {
    display: none; }
  
  /* line 3, (__TEMPLATE__) */
  .relative-time {
    display: auto; }
</style>
<style>
  /* line 1, (__TEMPLATE__) */
  .add_related_items {
    display: none; }
  
  /* line 3, (__TEMPLATE__) */
  #related_items {
    min-height: 75px; }
    /* line 5, (__TEMPLATE__) */
    #related_items .error {
      color: red; }
</style>
<link rel="stylesheet" media="screen" href="https://internal-cdn.amazon.com/code.amazon.com/pub/assets/cdn/spiffy_diffy_assets/spiffy_diffy-d3369a20c1b52ab59f433e45843299f2.css" />
<link rel="stylesheet" media="screen" href="https://internal-cdn.amazon.com/code.amazon.com/pub/assets/cdn/blobs-d793b26e66a0641d0db9e3b45482a842.css" />

</head>
<body>
<nav class='navbar navbar-default hidden-print' role='navigation'>
<div class='container-fluid'>
<div class='navbar-header'>
<a class='navbar-brand' href='/'>Code Browser</a>
</div>
<ul class='nav navbar-nav'>
<li>
<form action='/packages/search' class='form-inline navbar-search navbar-form'>
<div class='input-append'>
<input accesskey='s' class='hinted input-medium autocomplete search-query form-control search' data-autocomplete-url='/packages/autocomplete_package_id?vs=true' id='package_id_autocomplete' name='term' placeholder='Search' size='40' title='Search' type='text'>
</div>
</form>
</li>
<li>
<a href="/permissions">Permissions</a>
</li>
<li>
<a href="/workspaces/wangting">Workspaces</a>
</li>
<li>
<a href="/version-sets">Version Sets</a>
</li>
<li>
<a data-target='#preferences_dialog' data-toggle='modal' id='preferences'>Preferences</a>
<div class='modal fade' id='preferences_dialog' role='dialog'>
<div class='modal-dialog modal-md'>
<div class='modal-content'>
<div class='modal-header'>
<button class='close' data-dismiss='modal' type='button'>
<i class='fa fa-times'></i>
</button>
<h4 class='modal-title'>User Preferences</h4>
</div>
<div class='modal-body'>
<div class='text-center'>
<i class='fa fa-spinner fa-spin'></i>
</div>
</div>
</div>
</div>
</div>
</li>
</ul>
<div class='pull-right' id='delight'></div>
<ul class='nav navbar-nav navbar-right'>
<li><a target="_blank" onclick="window.open('http://tiny/submit/url?opaque=1&name=' + encodeURIComponent('/packages/PyAmazonCACerts/blobs/69d0fff03b4836a83bc2d1f6b74fd666eed7637b/--/README.md'+location.hash)); return false;" href="#">Tiny Link  <i class="fa fa-external-link"></i></a></li>
</ul>
<div id='query_string_instruction' style='display: none' title='Advanced Search'>
<h3>
Description
</h3>
<p>
Advanced Search allows user to specify search fields for each term. It also supports boolean algebra, wildcard and regular expressions.
<br>
If you are searching for phrases or sentences, Advanced Search should be able to provide more relevant result.
</p>
<h3>
Usage
</h3>
<ul>
<li>
Terms are seperated by spaces. Phrases or sentences should be quoted.
</li>
<li>
Wildcards like "?" or "*" are alllowed. But try to use leading wildcard with caution for better performance.
</li>
<li>
Regular expressions are also supported. To do that the string should be enclosed in two forward slash "/".
</li>
<li>
Specify a field name with "field_name:term". If a field name is not specified, the search will be done in "package" field.
</li>
<li>
Join multiple queries with "AND", "OR", and group them with "()". Be sure to type them in uppercase.
</li>
<li>
Fuzzy search is done by appending "~fuzziness_level" to the search term.
</li>
</ul>
<h3>
Fields
</h3>
<ul>
<li>
package (Brazil package name)
</li>
<li>
branch (package branch name)
</li>
<li>
file (file name)
</li>
<li>
path (file path)
</li>
<li>
code (source file content)
</li>
<li>
abstract (package short description)
</li>
<li>
description (package description)
</li>
</ul>
<h3>
Examples
</h3>
<table class='example_table'>
<tr class='example_table'>
<th class='example_table'>
Search for...
</th>
<th class='example_table'>
with query...
</th>
</tr>
<tr class='example_table'>
<td class='example_table'>
Package name you are not sure
</td>
<td class='example_table'>
AmazonABCWebsite~3
</td>
</tr>
<tr class='example_table'>
<td class='example_table'>
Sentence
</td>
<td class='example_table'>
description:"package created by octane"
</td>
</tr>
<tr class='example_table'>
<td class='example_table'>
Available clients for a service
</td>
<td class='example_table'>
package:ElasticSearchService*Client
</td>
</tr>
<tr class='example_table'>
<td class='example_table'>
Certain file type in a package
</td>
<td class='example_table'>
package:CodeBrowserWebsite AND file:*.rb
</td>
</tr>
<tr class='example_table'>
<td class='example_table'>
Package with given dependencies
</td>
<td class='example_table'>
file:Config AND (code:"Ruby-aws-sdk" AND code:ServiceMonitoring)
</td>
</tr>
</table>
<h3>
Reference
</h3>
<a href='https://www.elastic.co/guide/en/elasticsearch/reference/current/query-dsl-query-string-query.html' target='_blank'>
"Query String Query" in ElasticSearch document
</a>
</div>

</div>
</nav>

<div class='container-fluid'>
<ol class='breadcrumb'>
<li>
<a href="/">Home</a>
</li>
<li><a href="/packages/PyAmazonCACerts">PyAmazonCACerts</a></li>
<li><a href="/packages/PyAmazonCACerts/trees/mainline/--/">mainline</a></li>
<li class='.active'>README.md</li>


</ol>
<div id='content'>
</div>

<!-- Always showing this error message, but with the "hidden" class.  The package_badges script fetches more info and conditionally displays this. -->
<div class="alert hidden master_vs_problem alert-warning"><button type="button" class="close" data-dismiss="alert"><i class="fa fa-times-circle"></i></button><i class="ocon-white-warning"></i>This package appears to have a badly configured master version set for one of its major versions (most likely incorrectly set to "live").
Please go to the <a href="/packages/PyAmazonCACerts/releases">Releases</a> page to view and modify the master version set.
You can also read more about why you might care about this state <a href="https://w.amazon.com/index.php/BuilderTools/Product/OmniGrok/UserGuide#What_Brazil_code_gets_indexed.3F">here</a>.
</div><div class='page-header'>
<h1>
<span>
PyAmazonCACerts
</span>
<div class='star' data-package='PyAmazonCACerts'></div>
<div class='badges'>
<span id='third_party' style='display: none;'>
<span class='label label-info'>Third Party Package</span>
</span>
</div>
<small class='hidden-print'>
<a class='powertip autoselect pull-right' data-powertip='brazil ws use -p PyAmazonCACerts' id='bw_use'>
<i class='glyphicon glyphicon-download-alt'></i>
</a>
</small>
<small>
<span class='clone subtext pull-right hidden-print'>
<form class='form-inline'>
Clone uri:
<input class='form-control input-sm' type='text' value='ssh://git.amazon.com/pkg/PyAmazonCACerts'>
</form>
</span>

</small>
<small>
<div class='pull-right hidden-print' id='code_search_box'>
<form class="form-inline" action="/search_redirector" accept-charset="UTF-8" method="get"><input name="utf8" type="hidden" value="&#x2713;" />
<div class='input-group search '><input type="text" name="search_term" id="search_term" placeholder="Search OmniGrok within this package" size="33" class="form-control input-sm" />
<span class='input-group-btn'><button class='btn' type='submit'>Go</button></span></div><input type="hidden" name="package" id="package" value="PyAmazonCACerts" />
<input type="hidden" name="path" id="path" value="README.md" />
</form>

</div>

</small>
</h1>
</div>
<div class='row'>
<div class='col-md-9'>
<ul class='nav nav-pills bottom-buffer-small hidden-print'>
<li class='active'><a href="/packages/PyAmazonCACerts">Source</a></li>
<li><a href="/packages/PyAmazonCACerts/logs">Commits</a></li>
<li><a href="https://devcentral.amazon.com/ac/brazil/directory/package/overview/PyAmazonCACerts">Overview</a></li>
<li><a href="/packages/PyAmazonCACerts/releases">Releases</a></li>
<li><a href="/packages/PyAmazonCACerts/metrics/69d0fff03b4836a83bc2d1f6b74fd666eed7637b">Metrics</a></li>
<li><a href="/packages/PyAmazonCACerts/permissions">Permissions</a></li>
<li><a href="/packages/PyAmazonCACerts/repo-info">Repository Info</a></li>
</ul>

</div>
<div class='col-md-3'>
<div id='branch_and_search_box'>
<div class='hidden-print' id='branch_dropdown'>
<label for="branches">Branches: </label>
<input id='branches' name='branches' type='hidden'>
</div>

</div>

</div>
</div>
<div class='last_commit panel panel-default top-buffer-small'>
<div class='last_commit_heading'>
Last Commit
<span class='subtext'>
(<a class="commit-see-more" href="#">see more</a>)
</span>
</div>
<div class='panel-body'>
<ul class='last-commit-summary list-unstyled list-inline'>
<li class='commiter'></li>
<a href="https://code.amazon.com/users/jeffe/activity">Jeff Edwards</a>
<li class='time'></li>
<span title='Committed on January 12, 2016 09:30:11 AM PST' class='relative-time hover_tooltip year_old'>11 months ago</span><span class='absolute-time hover_tooltip year_old'>2016-01-12 09:30:11</span>
<li class='commit_message'>
<span class='refs'>
</span>
<a class='powertip commit black' data-commit-id='69d0fff03b4836a83bc2d1f6b74fd666eed7637b' href='/packages/PyAmazonCACerts/commits/69d0fff03b4836a83bc2d1f6b74fd666eed7637b'>
Add README for basic usage
</a>
</li>
<li><a class="mono powertip autoselect" data-powertip="69d0fff03b4836a83bc2d1f6b74fd666eed7637b" href="/packages/PyAmazonCACerts/commits/69d0fff03b4836a83bc2d1f6b74fd666eed7637b#README.md">69d0fff0</a></li>
<li>
<img src="https://pipelines.amazon.com/favicon.ico" alt="Favicon" />
<a href="https://pipelines.amazon.com/changes/PKG/PyAmazonCACerts/mainline/GitFarm:69d0fff03b4836a83bc2d1f6b74fd666eed7637b">Track in pipelines</a>
</li>
</ul>

<div class='swappable-with-brief-header'>
<div class='commit_header'>
<div class='portrait'><a href="https://code.amazon.com/users/jeffe/activity"><img class="" width="50" onerror="this.onerror=null; this.src=&#39;https://internal-cdn.amazon.com/code.amazon.com/pub/assets/cdn/default-user-39906d97be8dc4ee8b930b03f3e1b8f6.gif&#39;" src="https://internal-cdn.amazon.com/badgephotos.amazon.com/phone-image.pl?uid=jeffe" alt="Phone image" /></a></div>
<div class='details'>
<div class='pull-right' id='track_pipeline_change' style='clear-right'>
<ul class='list-unstyled'>
<li>
<img src="https://pipelines.amazon.com/favicon.ico" alt="Favicon" />
<a href="https://pipelines.amazon.com/changes/PKG/PyAmazonCACerts/mainline/GitFarm:69d0fff03b4836a83bc2d1f6b74fd666eed7637b">Track in pipelines</a>
<span class='subtext'>(mainline)</span>
</li>
</ul>
</div>
<div class='pull-right' id='browse_source' style='clear: right'>
<a href="/packages/PyAmazonCACerts/trees/69d0fff03b4836a83bc2d1f6b74fd666eed7637b">Browse source at this commit</a>
</div>
<div class='pull-right' id='child_link' style='clear: right'>
<a href="/packages/PyAmazonCACerts/commits/69d0fff03b4836a83bc2d1f6b74fd666eed7637b.child">view child commit</a>
</div>
<ul class='list-unstyled pull-right' style='clear: right'>
</ul>
<div class='author'>
<span class='name'><a href="https://code.amazon.com/users/jeffe/activity">Jeff Edwards</a></span>
<span class='sha1'>
(<a class='powertip autoselect' data-powertip='69d0fff03b4836a83bc2d1f6b74fd666eed7637b' href='/packages/PyAmazonCACerts/commits/69d0fff03b4836a83bc2d1f6b74fd666eed7637b'>69d0fff0</a>)
</span>
<div class='subtext'>
authored: <span title='January 12, 2016 09:30:11 AM PST' class='relative-time hover_tooltip year_old'>11 months ago</span><span class='absolute-time hover_tooltip year_old'>2016-01-12 09:30:11</span>, committed: <span title='January 12, 2016 09:30:11 AM PST' class='relative-time hover_tooltip year_old'>11 months ago</span><span class='absolute-time hover_tooltip year_old'>2016-01-12 09:30:11</span>
<div class='summaries'>
<div class='summary'>
Pushed to
<span class='autoselect branch powertip ref' data-powertip='mainline'>mainline</span>
by chrisros <span title='May 18, 2016 11:04:16 AM PDT' class='relative-time hover_tooltip year_old'>7 months ago</span><span class='absolute-time hover_tooltip year_old'>2016-05-18 11:04:16</span> as part of <a class='powertip autoselect' data-powertip='367f564733834fe3a3d57e5cfa9ea358dc286897' href='/packages/PyAmazonCACerts/commits/367f564733834fe3a3d57e5cfa9ea358dc286897'>367f5647</a>
</div>
</div>


</div>
<p class='top-buffer'>
<span class='subject'>
<a href="/packages/PyAmazonCACerts/commits/69d0fff03b4836a83bc2d1f6b74fd666eed7637b">Add README for basic usage</a>
</span>
</p>
</div>
<div id='related_items'>
<h3>Related Items</h3>
<div class='fetching subtext'>
Fetching...
</div>
<div class='msg subtext' style='display: none'>
No related items found.
</div>
<ul data-bind='foreach: relatedItemsModel().relatedItems, visible: relatedItemsModel().relatedItems().length &gt; 0'>
<li>
<span data-bind='text: type'></span>
<a data-bind='text: link.title, attr: {href: link.url}'></a>
<a class='delete_relation' data-bind="attr: {href: '/delete-relation?eid=' + link.eid}" onclick='return confirm("Really delete this relation?")'>
<span class='red'>&nbsp;x&nbsp;</span>
</a>
</li>
</ul>
<div class='add_relation_link'>
<a href='#'>+ Add Relation</a>
</div>
<div class='add_related_items'>
<form action="/create_relation" accept-charset="UTF-8" method="post"><input name="utf8" type="hidden" value="&#x2713;" /><input type="hidden" name="authenticity_token" value="Dw1qoQTL0Q1dIywCmaNjuu7bP7OhgZnODJKbpOnvzc/+t8boriu1Amja1U4xGF4F62mSkDkyLp4pSAMri+PqEg==" />
Relate this commit to url:
<input name='relation' type='text'>
<input type="hidden" name="package_id" id="package_id" value="PyAmazonCACerts" />
<input type="hidden" name="commit_id" id="commit_id" value="69d0fff03b4836a83bc2d1f6b74fd666eed7637b" />
<input type="submit" name="commit" value="Save" class="btn btn-default" />
</form>

</div>
</div>
<div class='pull-right' id='create_branch'>
<a href='#'>Create branch</a>
<form action="/packages/PyAmazonCACerts/create-branch" accept-charset="UTF-8" method="post"><input name="utf8" type="hidden" value="&#x2713;" /><input type="hidden" name="authenticity_token" value="QuqfGPfPelV4Vva59162oFs/QIHgJHyX9KHgB3aDeluzUDNRXS8eWk2vD/Vf5YsfXo3toniXy8fRe3iIFI9dhg==" />
Create branch named
<input type="text" name="branch_name" id="branch_name" />
from this commit
<input type="hidden" name="source_sha1" id="source_sha1" value="69d0fff03b4836a83bc2d1f6b74fd666eed7637b" />
<input type="submit" name="commit" value="Go!" class="btn btn-default" />
</form>

</div>
</div>
</div>

</div>
</div>
</div>
<div class='clear'></div>

<div class='jump_to_file hidden-print'>
<div class='jump_to_file_form'>
<form class='form_inline' onSubmit='return false'>
<input type="hidden" name="package_id" id="package_id" value="PyAmazonCACerts" />
<input type="hidden" name="commit_id_for_file" id="commit_id_for_file" value="mainline" />
<div class='input-append'>
<input accesskey='j' class='hinted form-control search' id='filesearch' name='file' placeholder='Jump to a file' title='Jump to a file' type='text'>
</div>
<div class='jump_to_file_dismiss'></div>
</form>
</div>
<div class='jump_to_file_popup'><a class='help helpPopup' data-content="Here you can enter the name of the file and it will provide suggestions with the matching file names and the path for the same.&lt;br/&gt; After selecting the required file, it will redirect to that file. The keyboard shortcut is 'CTRL+j'.">
<img src='https://internal-cdn.amazon.com/btk.amazon.com/img/icons-1.0/tooltip-bubble.png'>
</a>
</div>
<div class='jump_to_file_error'>
The above file can not be found. Either the whole path is missing or the file is not in
<br>this package. Please check the autosuggestions.</br>
</div>
</div>

<!--
mime_type: text/plain
-->
<div class='file_header'>
<div class='path_breadcrumbs'>
<div class='path_breadcrumbs'>
<span class='path_breadcrumb'><a href="/packages/PyAmazonCACerts">PyAmazonCACerts</a></span> / <span class='path_breadcrumb'><a href="/packages/PyAmazonCACerts/trees/mainline/--/">mainline</a></span> / <span class='path_breadcrumb'>README.md</span>
</div>

</div>
<div class='hidden-print' id='file_actions'>
<ul class='button_group'>
<li>
<a class="minibutton" href="/packages/PyAmazonCACerts/blobs/mainline/--/README.md?raw=1">Raw</a>
</li>
<li>
<a class="minibutton" href="/packages/PyAmazonCACerts/blobs/mainline/--/README.md?download=1">Download</a>
</li>
<li>
<a class="minibutton" href="/packages/PyAmazonCACerts/logs/mainline?path=README.md">History</a>
</li>
<li>
<a class="minibutton" href="/packages/PyAmazonCACerts/blobs/mainline/--/README.md/edit_file_online">Edit</a>
</li>
<li>
<a class="minibutton md_control_raw" href="#raw_markdown">Markdown Source</a>
</li>
<li>
<a class="minibutton md_control_rendered selected active" href="#">Rendered Markdown</a>
</li>
<li class='permalink'>
<a class="minibutton" href="/packages/PyAmazonCACerts/blobs/69d0fff03b4836a83bc2d1f6b74fd666eed7637b/--/README.md">Permalink</a>
</li>
</ul>
</div>
<div class='clear'></div>
<markdown add-raw-if-needed='true'>
## PyAmazonCACerts

If you're using requests for your https calls (_and you should_) simply adding this to your dependencies/runtime-dependencies (depending on where you're calling it) and importing it at the top of your file will put all the right stuff together for you, giving you a merged cert verification file against both internal and externally-signed CAs.

import amazoncerts

The module is super simple (&lt;10 lines as of writing this).

</markdown>
<div class='blob hidden-print highlighttable markdown_source' ng_non_bindable>
    <div class="js-syntax-highlight-wrapper">
      <table class="code js-syntax-highlight">
        <tbody>
            <tr class="line_holder" id="L1">
              <td class="line-num" data-linenumber="1">
                <span class="linked-line" unselectable="on" data-linenumber="1"></span>
              </td>
              <td class="line_content"><span class="gu">## PyAmazonCACerts</span>
</td>
            </tr>
            <tr class="line_holder" id="L2">
              <td class="line-num" data-linenumber="2">
                <span class="linked-line" unselectable="on" data-linenumber="2"></span>
              </td>
              <td class="line_content">
</td>
            </tr>
            <tr class="line_holder" id="L3">
              <td class="line-num" data-linenumber="3">
                <span class="linked-line" unselectable="on" data-linenumber="3"></span>
              </td>
              <td class="line_content">If you&#39;re using requests for your https calls (_and you should_) simply adding this to your dependencies/runtime-dependencies (depending on where you&#39;re calling it) and importing it at the top of your file will put all the right stuff together for you, giving you a merged cert verification file against both internal and externally-signed CAs.
</td>
            </tr>
            <tr class="line_holder" id="L4">
              <td class="line-num" data-linenumber="4">
                <span class="linked-line" unselectable="on" data-linenumber="4"></span>
              </td>
              <td class="line_content">
</td>
            </tr>
            <tr class="line_holder" id="L5">
              <td class="line-num" data-linenumber="5">
                <span class="linked-line" unselectable="on" data-linenumber="5"></span>
              </td>
              <td class="line_content">import amazoncerts
</td>
            </tr>
            <tr class="line_holder" id="L6">
              <td class="line-num" data-linenumber="6">
                <span class="linked-line" unselectable="on" data-linenumber="6"></span>
              </td>
              <td class="line_content">
</td>
            </tr>
            <tr class="line_holder" id="L7">
              <td class="line-num" data-linenumber="7">
                <span class="linked-line" unselectable="on" data-linenumber="7"></span>
              </td>
              <td class="line_content">The module is super simple (&lt;10 lines as of writing this).</td>
            </tr>
        </tbody>
      </table>
    </div>

</div>
</div>


</div>
<nav class='navbar navbar-default footer' role='navigation'>
<footer class='footer top-buffer' id='footer'>
<div class='col-sm-9 col-md-8 main'>
<h3>Packages</h3>
<ul class='unstyled'>
<li><a href="https://octane.amazon.com/package">Create Package</a></li>
<li><a href="/packages/find_by_team_for_user">All packages for my team</a></li>
</ul>
<h3>Commit Notifications</h3>
<ul class='unstyled'>
<li><a href="https://w.amazon.com/index.php/BuilderTools/Product/RevisionControl/CommitNotifications">RSS</a></li>
<li><a href="/commit-notifications">Email</a></li>
</ul>
</div>
<div class='col-sm-3 col-md-4 sidebar'>
<div class='business_card clearfix'>
<h3>Need help?</h3>
<ul class='unstyled'>
<li><a href="https://issues.amazon.com/issues/create?assignedFolder=9bbc2895-0c1a-4981-9a45-d81dbbdb683e&amp;description=***+Please+-+fill+in+the+fields+below%2C+so+we+can+better+assist+you.++Thanks.%0A%0A%0AI+was+using+the+Code+Browser+and+[encountered+a+problem+%2F+have+a+suggestion].%0A%0AThe+URL+of+the+page+I+was+viewing%3A+[paste+URL%2C+or+%22N%2FA%22]%0A%0AWhat+I+did%3A+[describe+action]%0A%0AWhat+I+wanted%3A+[describe+expected+behavior%2C+or+desired+new+behavior]%0A%0AWhat+actually+happened%3A+[describe+observed+behavior]%0A%0AAnything+else+relevant%3A+[...]&amp;descriptionContentType=text%2Fplain&amp;assigneeIdentity=">Submit an Issue (problems or suggestions)</a></li>
<li><a href="https://w.amazon.com/index.php/BuilderTools/Product/CodeBrowser">Code Browser Documentation</a></li>
<li><a href="https://w.amazon.com/?DTUX/Browser_Support_Policy">Browser Support Policy</a></li>
</ul>
</div>
</div>
</footer>
</nav>

<script>
  var codeBrowserSpoofedUser = "wangting"
</script>
<script src="https://internal-cdn.amazon.com/code.amazon.com/pub/assets/cdn/vendor-8ce2e42e17e7eecfffbbc8f95ce77f6b.js"></script>
<script src="https://internal-cdn.amazon.com/code.amazon.com/pub/assets/cdn/application-9019950bfb39501df1699a4dbd66deef.js"></script>
<script src="https://internal-cdn.amazon.com/is-it-down.amazon.com/javascripts/stripe.min.js"></script>
<script>
  (function() {
    $(function() {
      return isItDownStripe('sourcecode', 1107, 1);
    });
  
  }).call(this);
</script>
<script src="https://internal-cdn.amazon.com/code.amazon.com/pub/assets/cdn/application_angular-225559ddd3da7b66ebf7938c90d01125.js"></script>
<script>
  bootstrapCodeBrowserNgApp();
</script>
<script>
  $(document).ready(function() {
      $('#branches').select2({
          width: "274px",
          data: [{"text":"Official Branches","children":[{"text":"mainline (default)","id":"/packages/PyAmazonCACerts/blobs/heads/mainline/--/README.md"}]},{"text":"calvinm's Shared Branches","children":[{"text":"rpm","id":"/packages/PyAmazonCACerts/blobs/share/calvinm/rpm/--/README.md"}]},{"text":"lorenzla's Shared Branches","children":[{"text":"mainline","id":"/packages/PyAmazonCACerts/blobs/share/lorenzla/mainline/--/README.md"}]},{"text":"mejiaa's Shared Branches","children":[{"text":"mainline","id":"/packages/PyAmazonCACerts/blobs/share/mejiaa/mainline/--/README.md"}]}],
          createSearchChoice: function(term, data) {
            if ($(data).filter(function() { return this.text.localeCompare(term)===0; }).length===0) {
              // This code fires if the user enters a string and hits return (rather than selecting the item
              // from the dropdown.  This breaks when viewing commits (logs). Customize it accordingly.
              if ('/packages/PyAmazonCACerts/blobs/'.match(/\/logs/)) {
                var id_string = '/packages/PyAmazonCACerts/blobs//' + term;
                if ('README.md') {
                  id_string += '/--/README.md';
                }
                return {text:term, id: id_string};
              }
              return {text:term, id:'/packages/PyAmazonCACerts/blobs/' + term + '/--/README.md'};
            }
          },
      });
      $('#branches').select2('data', null);
      $('#branches').change(function() {
        document.location = $(this).val();
      });
  });
</script>
<script>
  (function() {
    $(function() {
      return $('.add_relation_link a').click(function() {
        $('.add_related_items').show(500).find('input[name=relation]').delay(500).focus();
        $(this).hide();
        return false;
      });
    });
  
  }).call(this);
</script>
<script>
  (function() {
    $(function() {
      return $('.commit-see-more').click(function() {
        $('.last-commit-summary, .commit_header').toggle();
        return false;
      });
    });
  
  }).call(this);
</script>
<script>
  (function() {
    $('li.permalink a').click(function() {
      window.location.href = this.href + location.hash;
      return false;
    });
  
  }).call(this);
</script>
<script>
  (function() {
    var relatedItems;
  
    relatedItems = new Codac.RelatedItemsModel("PyAmazonCACerts", "69d0fff03b4836a83bc2d1f6b74fd666eed7637b", 'mainline', '');
  
    Codac.model.relatedItemsModel(relatedItems);
  
  }).call(this);
</script>
<script src="https://internal-cdn.amazon.com/code.amazon.com/pub/assets/cdn/spiffy_diffy_assets/spiffy_diffy-f11e935e37ccf6a6263d8da191fef501.js"></script>
<script>
  (function() {
    var key, onUrlChange, premalinkBtnEl, premalinkPath, template;
  
    (function() {
      var anchor, anchorMatch, hash, hl_lines_match, path, ranges, search;
      anchor = window.location.hash.split('#')[1] || '';
      anchorMatch = anchor.match(/^line-(\d+)/);
      if (anchorMatch) {
        anchor = 'L' + anchorMatch[1];
      }
      search = window.location.search;
      hl_lines_match = search.match(/hl_lines=([\d\-\,]+)/);
      ranges = '';
      if (hl_lines_match) {
        ranges = hl_lines_match[1].split(',').map(function(range) {
          return range.split('-').map(function(lineNum) {
            return 'L' + lineNum;
          }).join('-');
        }).join(',');
      }
      hash = ranges;
      if (anchorMatch) {
        hash += '|' + anchor;
      }
      if (hash !== '') {
        if (hash !== '') {
          window.location.hash = '#' + hash;
        }
        path = window.location.pathname + window.location.hash;
        return window.history.pushState(void 0, void 0, path);
      }
    })();
  
    premalinkBtnEl = $('#file_actions .permalink');
  
    premalinkPath = premalinkBtnEl.find('.minibutton').attr('href');
  
    onUrlChange = function() {
      return premalinkBtnEl.hide();
    };
  
    setTimeout((function() {
      return (new Diff()).enableHighlighting({
        basePath: premalinkPath,
        onUrlChange: onUrlChange
      });
    }), 0);
  
    key = 'codac:blob:line_highlighting';
  
    template = _.template("<p><strong>Line highlighting functionality was updated!</strong></p><p>For more information see our <a href='https://w.amazon.com/index.php/BuilderTools/Product/CodeBrowser/UserGuide#How_can_I_highlight_a_block_of_code.3F'>wiki page</a>.</p>");
  
    Codac.WhatsNew.create({
      key: key,
      template: template
    });
  
  }).call(this);
</script>
<script>
  $(document).ready(function(){
    // print the query in search box
    // use @query.inspect to get the query string, since
    // the backslash passed by controller is not escaped
    // if use @query directily as text in js, the backslash
    // will fail to display
    $("#code_search_input").val("");
    if("" == "on")
      $("#use_query_string").prop('checked', true);
    // refill input
    $("#code_search_input").focusout(function(){
      if($("#code_search_input").val() == '')
        $("#code_search_input").val("");
    });
    // clear input button
    $("#clear-input-button").click(function(){
      $("#code_search_input").val('');
      $("#code_search_input").focus();
    });
    // make outer div also clickable
    $("#use_query_string_checkbox").click(function(){
      $("#use_query_string").prop('checked',!$("#use_query_string").prop('checked'));
    })
  });
</script>

</body>
</html>
