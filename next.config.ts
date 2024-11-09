import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  output: 'export',
  basePath: '/webercformatter',
  images: {
    unoptimized: true
  },
};

module.exports = nextConfig;
