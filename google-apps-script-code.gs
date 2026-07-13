function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
  var params = e.parameter;
  
  if (params.password !== 'inksecret' && params.action !== 'add_comment' && params.action !== 'add_like') {
    return ContentService.createTextOutput(JSON.stringify({success: false, error: 'Unauthorized'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var action = params.action; // 'update', 'delete', or undefined (for new post/thought)
  var slug = params.slug;
  var target = params.target || 'home';

  if (action === 'add_comment') {
    var commentSheet = ss.getSheetByName('Comments') || ss.insertSheet('Comments');
    commentSheet.appendRow([target, params.name || 'Anonymous', params.comment || '', new Date().toISOString()]);
    return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'add_like') {
    var likeSheet = ss.getSheetByName('Likes') || ss.insertSheet('Likes');
    var likeData = likeSheet.getDataRange().getValues();
    var found = false;
    for (var i = 0; i < likeData.length; i++) {
      if (likeData[i][0] === target) {
        likeSheet.getRange(i + 1, 2).setValue((Number(likeData[i][1]) || 0) + 1);
        found = true;
        break;
      }
    }
    if (!found) {
      likeSheet.appendRow([target, 1]);
    }
    return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'delete') {
    if (!slug) return ContentService.createTextOutput(JSON.stringify({success: false, error: 'No slug provided'})).setMimeType(ContentService.MimeType.JSON);
    
    var data = sheet.getDataRange().getValues();
    // Loop from bottom up to avoid index shifting on delete
    // Start at data.length - 1 (not skipping row 0, since there is no header row in Sheet1)
    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]).trim() === String(slug).trim()) { 
        sheet.deleteRow(i + 1); // +1 because array is 0-indexed, sheet rows are 1-indexed
        return ContentService.createTextOutput(JSON.stringify({success: true, message: 'Deleted'}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({success: false, error: 'Post not found'})).setMimeType(ContentService.MimeType.JSON);
  } 
  
  if (action === 'update') {
    if (!slug) return ContentService.createTextOutput(JSON.stringify({success: false, error: 'No slug provided'})).setMimeType(ContentService.MimeType.JSON);
    
    var data = sheet.getDataRange().getValues();
    // Start at i=0 — no header row in Sheet1
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(slug).trim()) {
        // Col A(1)=slug, B(2)=title, C(3)=excerpt, D(4)=tags, E(5)=image, F(6)=date, G(7)=author, H(8)=readTime, I(9)=content
        sheet.getRange(i + 1, 2).setValue(params.title || '');
        sheet.getRange(i + 1, 3).setValue(params.excerpt || '');
        sheet.getRange(i + 1, 4).setValue(params.tags || '');
        sheet.getRange(i + 1, 5).setValue(params.image || '');
        sheet.getRange(i + 1, 8).setValue(params.readTime || '');
        sheet.getRange(i + 1, 9).setValue(params.content || '');
        
        return ContentService.createTextOutput(JSON.stringify({success: true, message: 'Updated'}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({success: false, error: 'Post not found'})).setMimeType(ContentService.MimeType.JSON);
  }

  // --- ORIGINAL ADD POST / UPDATE THOUGHT LOGIC ---
  var rowData = [];
  rowData.push(params.slug || '');
  rowData.push(params.title || '');
  rowData.push(params.excerpt || '');
  rowData.push(params.tags || '');
  rowData.push(params.image || '');
  rowData.push(params.date || '');
  rowData.push(params.author || '');
  rowData.push(params.readTime || '');
  rowData.push(params.content || '');
  rowData.push(params.link || '');
  rowData.push(params.thought || '');
  
  sheet.appendRow(rowData);
  
  return ContentService.createTextOutput(JSON.stringify({success: true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var params = e.parameter;
  var result = {};
  var target = params.target || 'home';
  
  if (params.action === 'get_comments') {
    var commentSheet = ss.getSheetByName('Comments');
    var comments = [];
    if (commentSheet) {
      var commentData = commentSheet.getDataRange().getValues();
      for (var i = 0; i < commentData.length; i++) {
        if (commentData[i][0] === target) {
          comments.push({
            name: String(commentData[i][1]),
            comment: String(commentData[i][2]),
            date: String(commentData[i][3])
          });
        }
      }
    }
    result = { comments: comments };
  } else if (params.action === 'get_all_likes') {
    // Returns all likes as { likes: { slug: count, ... } }
    var likeSheet = ss.getSheetByName('Likes');
    var allLikes = {};
    if (likeSheet) {
      var likeData = likeSheet.getDataRange().getValues();
      for (var i = 0; i < likeData.length; i++) {
        var key = String(likeData[i][0]).trim();
        if (key) {
          allLikes[key] = Number(likeData[i][1]) || 0;
        }
      }
    }
    result = { likes: allLikes };
  } else if (params.action === 'get_likes') {
    var likeSheet = ss.getSheetByName('Likes');
    var likes = 0;
    if (likeSheet) {
      var likeData = likeSheet.getDataRange().getValues();
      for (var i = 0; i < likeData.length; i++) {
        if (likeData[i][0] === target) {
          likes = Number(likeData[i][1]) || 0;
          break;
        }
      }
    }
    result = { likes: likes };
  } else if (params.action === 'thought') {
    // Return last thought (original logic)
    for (var i = data.length - 1; i >= 1; i--) {
      var thought = data[i][10]; // assuming thought is in Col K (index 10)
      if (thought && String(thought).trim() !== '') {
        result = { thought: String(thought).trim() };
        break;
      }
    }
  } else if (params.action === 'posts' || !params.action || params.action === 'post') {
    // Return posts or single post
    
    // First, build a map of all likes so we can include it in the posts
    var likeSheet = ss.getSheetByName('Likes');
    var likesMap = {};
    if (likeSheet) {
      var likeData = likeSheet.getDataRange().getValues();
      for (var l = 0; l < likeData.length; l++) {
        likesMap[likeData[l][0]] = Number(likeData[l][1]) || 0;
      }
    }

    if (params.action === 'post' && params.slug) {
        var post = null;
        for (var i = 1; i < data.length; i++) {
          if (data[i][0] === params.slug) {
             post = {
                slug: data[i][0],
                title: String(data[i][1]),
                excerpt: String(data[i][2] || ''),
                tags: String(data[i][3] || ''),
                image: String(data[i][4] || ''),
                date: String(data[i][5] || ''),
                author: String(data[i][6] || ''),
                readTime: String(data[i][7] || ''),
                content: String(data[i][8] || ''),
                likes: likesMap[data[i][0]] || 0
             };
             break;
          }
        }
        result = { post: post };
    } else {
      var posts = [];
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        // Only include if there's a title and it's not empty
        if (row[1] && String(row[1]).trim() !== '') {
          posts.push({
            slug: row[0],
            title: String(row[1]),
            excerpt: String(row[2] || ''),
            tags: String(row[3] || ''),
            image: String(row[4] || ''),
            date: String(row[5] || ''),
            author: String(row[6] || ''),
            readTime: String(row[7] || ''),
            content: String(row[8] || ''),
            likes: likesMap[row[0]] || 0
          });
        }
      }
      result = { posts: posts };
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
