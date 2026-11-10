const {app,BrowserWindow,dialog,shell}=require('electron');const path=require('node:path');const fs=require('node:fs');const crypto=require('node:crypto');let httpServer,store;
if(!app.requestSingleInstanceLock())app.quit();
app.on('second-instance',()=>{const w=BrowserWindow.getAllWindows()[0];if(w){if(w.isMinimized())w.restore();w.focus();}});
app.whenReady().then(async()=>{
 try{
  const folder=app.getPath('userData');const configPath=path.join(folder,'config.json');fs.mkdirSync(folder,{recursive:true});
  let config;if(fs.existsSync(configPath))config=JSON.parse(fs.readFileSync(configPath,'utf8'));else{config={secret:crypto.randomBytes(48).toString('hex'),setupToken:crypto.randomBytes(12).toString('hex')};fs.writeFileSync(configPath,JSON.stringify(config),{mode:0o600});}
  process.env.DB_KIND='sqlite';process.env.SQLITE_PATH=path.join(folder,'courses.sqlite');process.env.SESSION_SECRET=config.secret;process.env.SETUP_TOKEN=config.setupToken;process.env.BACKUP_DIR=path.join(folder,'backups');
  const {createApp}=require('../src/server');const result=await createApp();store=result.store;
  httpServer=await new Promise(resolve=>{const server=result.app.listen(0,'127.0.0.1',()=>resolve(server));});const origin=`http://127.0.0.1:${httpServer.address().port}`;
  const win=new BrowserWindow({width:1360,height:900,minWidth:850,minHeight:650,title:'همراه محصل — پوهنتون کندوز',autoHideMenuBar:true,webPreferences:{nodeIntegration:false,contextIsolation:true,sandbox:true}});
  win.webContents.on('will-navigate',(event,url)=>{if(new URL(url).origin!==origin)event.preventDefault();});
  win.webContents.setWindowOpenHandler(({url})=>{if(url.startsWith(origin+'/api/certificates/print'))return {action:'allow',overrideBrowserWindowOptions:{autoHideMenuBar:true,webPreferences:{nodeIntegration:false,contextIsolation:true,sandbox:true}}};if(url==='about:blank')return {action:'allow',overrideBrowserWindowOptions:{webPreferences:{nodeIntegration:false,contextIsolation:true,sandbox:true}}};return {action:'deny'};});
  win.webContents.session.setPermissionRequestHandler((_wc,_permission,callback)=>callback(false));
  win.webContents.session.on('will-download',(_e,item)=>{item.setSaveDialogOptions({defaultPath:path.join(app.getPath('downloads'),item.getFilename())});});
  await win.loadURL(origin);
  if(!(await store.load()).users.length)await dialog.showMessageBox(win,{type:'info',title:'راه‌اندازی نخستین مدیر',message:'برای ساخت حساب مدیر، کد زیر را در صفحه وارد کنید:',detail:config.setupToken+'\n\nکد در فایل config.json پوشهٔ اطلاعات برنامه نیز ذخیره شده است.'});
 }catch(e){dialog.showErrorBox('خطا در اجرای برنامه',e.message);app.quit();}
});
app.on('window-all-closed',()=>app.quit());app.on('before-quit',()=>{if(httpServer)httpServer.close();if(store?.db){try{store.db.close();}catch{}}});
