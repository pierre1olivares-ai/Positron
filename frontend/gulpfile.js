'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);
// Q-Star's generated Tailwind utilities are intentionally global but scoped beneath .qstar-app.
build.addSuppression(`Warning - [sass] src/webparts/qstarIssueManager/components/QstarPrototype.scss: filename should end with module.sass or module.scss`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
