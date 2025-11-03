import { Config } from './shared/types';

// Read configuration from environment variables (provided via dotenv-webpack)
const config: Config = {
  baseRoot: process.env.BASE_ROOT || 'https://api.your-domain.com',
  token: process.env.TOKEN || undefined
};
export { config };
