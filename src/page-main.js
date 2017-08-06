console.log('aaa')
import './page-main.scss';
import OutlookUtil from './utils/OutlookUtil';


OutlookUtil.initialize()
    .then((reason) => {
        console.log('BOOM...', reason);
    });

console.log('bbb');
