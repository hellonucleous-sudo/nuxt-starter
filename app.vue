<script setup>
useHead({
  htmlAttrs: { lang: "en-US" },
  title: "Nucleous",
  meta: [
    {
      name: "description",
      content:
        "Nucleous is an AI-driven CRM and workforce intelligence platform that helps businesses manage leads, funnels, teams, follow-ups, quotations, and analytics in one unified system."
    },
    {
      name: "keywords",
      content:
        "AI CRM, lead management software, sales funnel CRM, workforce management platform, AI follow-ups, quotation builder, sales automation"
    },
    {
      property: "og:title",
      content: "Nucleous | AI-Driven CRM & Workforce Intelligence Platform"
    },
    {
      property: "og:description",
      content:
        "Manage leads, pipelines, teams, AI follow-ups, and analytics from a single intelligent platform. From cold leads to closed deals, Nucleous keeps everything connected."
    },
    {
      property: "og:type",
      content: "website"
    },
    {
      property: "og:url",
      content: "https://nucleous.io"
    },
    {
      property: "og:image",
      content: "https://nucleous.io/og-image.jpg"
    },
    {
      name: "twitter:card",
      content: "summary_large_image"
    },
    {
      name: "twitter:title",
      content: "Nucleous | AI-Driven CRM & Workforce Intelligence Platform"
    },
    {
      name: "twitter:description",
      content:
        "AI-powered CRM to manage leads, funnels, teams, follow-ups, and analytics with enterprise-grade control."
    }
  ],
  bodyAttrs: {
    class: "font-display antialiased bg-white text-gray-900"
  }
});
</script>

<template>
  <NuxtLayout>
    <NuxtPage />
  </NuxtLayout>
</template>
