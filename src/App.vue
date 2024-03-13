<template>
  <v-app>
    <v-main>
      <Wages/>
    </v-main>
  </v-app>
</template>

<script>
import Wages from './components/Wages';

export default {
  name: 'App',

  components: {
    Wages,
  },

  data: () => ({
    //
  }),
};
</script>
