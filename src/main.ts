import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';

//platformBrowserDynamic().bootstrapModule(AppModule);

// For running within Excel
Office.initialize = function () {
  platformBrowserDynamic().bootstrapModule(AppModule);
};
