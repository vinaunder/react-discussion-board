import { sp } from '@pnp/sp';
import '@pnp/sp/site-users/web';
export const useUsers = () => {
  // Run on useList hook
  (async () => {})();

  const currentUser = async(): Promise<any> => {
    let user = await sp.web.currentUser.get();
    return user;
  };
  // Return functions
  return {
    currentUser,
  };
};
