// Quick test script
import fs from 'fs';

const testPath = '\\\\172.17.1.3\\cylande\\TomcatUR\\UnitedRetailAziza\\Data\\default\\In\\';

try {
  const files = fs.readdirSync(testPath);
  console.log('✅ Path accessible!');
  console.log('Files found:', files.length);
} catch (error) {
  console.error('❌ Cannot access path:', error.message);
}